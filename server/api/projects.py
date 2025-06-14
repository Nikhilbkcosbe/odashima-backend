import uuid
import traceback
from datetime import datetime
from fastapi import APIRouter, HTTPException, status, Depends
from fastapi.responses import JSONResponse
from server.helpers.auth import OAuth2PasswordBearerWithCookie
from server.configs.db import projects_collection
from pydantic import BaseModel
from typing import Optional

# Create a new router
router = APIRouter()

# Define the OAuth2 scheme for authentication
oauth2_scheme = OAuth2PasswordBearerWithCookie(tokenUrl="/api/v1/auth/login")

# Define the project data model


class ProjectInputDataModel(BaseModel):
    project_name: str
    description: Optional[str] = None

# Create a new project


@router.post("/projects/create")
async def create_project(project_data: ProjectInputDataModel, current_user: str = Depends(oauth2_scheme)):
    """Create a new project.

    Args:
        project_data (ProjectInputDataModel): The project's data including name, description, start date, and end date.
        current_user (str): The current authenticated user.

    Returns:
        JSONResponse: A response indicating the creation status.

    Raises:
        HTTPException: If the user is not authorized or an error occurs during the process.
    """
    try:
        # Create a new project document
        new_project = {
            "_id": str(uuid.uuid4()),
            "project_name": project_data.project_name,
            "description": project_data.description,
            "created_at": datetime.now(),
            "created_by": current_user["email"]
        }

        # Insert the project into the database
        await projects_collection.insert_one(new_project)

        # Get all projects for the user
        projects = await projects_collection.find(
            {"created_by": current_user["email"]}
        ).to_list(length=None)

        # Format dates for response
        formatted_projects = []
        for project in projects:
            formatted_project = project.copy()
            if "created_at" in formatted_project and formatted_project["created_at"]:
                formatted_project["created_at"] = formatted_project["created_at"].isoformat(
                )
            if "updated_at" in formatted_project and formatted_project["updated_at"]:
                formatted_project["updated_at"] = formatted_project["updated_at"].isoformat(
                )
            formatted_projects.append(formatted_project)

        content = {"message": "Project created successfully",
                   "projects": formatted_projects}
        return JSONResponse(status_code=status.HTTP_201_CREATED, content=content)

    except HTTPException as e:
        raise e
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(e)
        ) from e

# Get all projects


@router.get("/projects")
async def get_all_projects(current_user: str = Depends(oauth2_scheme)):
    """Get all projects for the current user.

    Args:
        current_user (str): The current authenticated user.

    Returns:
        JSONResponse: A response containing the list of projects.

    Raises:
        HTTPException: If the user is not authorized.
    """
    try:

        # Retrieve the user's projects from the database
        projects = await projects_collection.find({}).to_list(length=None)

        # Format dates for response
        formatted_projects = []
        for project in projects:
            formatted_project = project.copy()
            if "created_at" in formatted_project and formatted_project["created_at"]:
                formatted_project["created_at"] = formatted_project["created_at"].isoformat(
                )
            if "updated_at" in formatted_project and formatted_project["updated_at"]:
                formatted_project["updated_at"] = formatted_project["updated_at"].isoformat(
                )
            formatted_projects.append(formatted_project)

        content = {"projects": formatted_projects}
        return JSONResponse(status_code=status.HTTP_200_OK, content=content)

    except HTTPException as e:
        raise e
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(e)
        ) from e

# Update a project


@router.put("/projects/{project_id}")
async def update_project(project_id: str, project_data: ProjectInputDataModel, current_user: str = Depends(oauth2_scheme)):
    """Update a project.

    Args:
        project_id (str): The ID of the project to update.
        project_data (ProjectInputDataModel): The updated project data.
        current_user (str): The current authenticated user.

    Returns:
        JSONResponse: A response indicating the update status.

    Raises:
        HTTPException: If the user is not authorized or the project is not found.
    """
    try:
        # Find the project to update
        project = await projects_collection.find_one({"_id": project_id})
        if not project:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Project not found.",
            )

        # Check if the user is the creator of the project
        if project["created_by"] != current_user["email"]:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="You do not have permission to update this project.",
            )

        # Update the project
        update_data = {
            "project_name": project_data.project_name,
            "description": project_data.description,
            "updated_at": datetime.now(),
            "updated_by": current_user["email"]
        }

        # Update the project document
        await projects_collection.update_one(
            {"_id": project_id},
            {"$set": update_data}
        )

        # Get the updated project
        updated_project = await projects_collection.find_one({"_id": project_id})

        # Format dates for response
        if "created_at" in updated_project and updated_project["created_at"]:
            updated_project["created_at"] = updated_project["created_at"].isoformat()
        if "updated_at" in updated_project and updated_project["updated_at"]:
            updated_project["updated_at"] = updated_project["updated_at"].isoformat()

        return JSONResponse(
            status_code=status.HTTP_200_OK,
            content={
                "message": "Project updated successfully",
                "project": updated_project
            }
        )

    except HTTPException as e:
        raise e
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=str(e)
        ) from e

# Delete a project


@router.delete("/projects/{project_id}")
async def delete_project(project_id: str, current_user: str = Depends(oauth2_scheme)):
    """Delete a project.

    Args:
        project_id (str): The ID of the project to delete.
        current_user (str): The current authenticated user.

    Returns:
        JSONResponse: A response indicating the deletion status.

    Raises:
        HTTPException: If the user is not authorized or the project is not found.
    """
    try:


        # Find the project to delete
        project = await projects_collection.find_one({"_id": project_id})
        if not project:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Project not found.",
            )

        # Check if the user is the creator of the project
        if project["created_by"] != current_user["email"]:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="You do not have permission to delete this project.",
            )

        # Delete the project
        await projects_collection.delete_one({"_id": project_id})

        return JSONResponse(
            status_code=status.HTTP_200_OK,
            content={
                "message": "Project deleted successfully"
            }
        )

    except HTTPException as e:
        raise e
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=str(e)
        ) from e
