import os
import traceback
from datetime import datetime, timedelta
import uuid
from fastapi import APIRouter, HTTPException, status, Response, Request
from server.schemas.login import LoginInputDataModel, ForgotPasswordInputDataModel, ResetPasswordInputDataModel
from dotenv import load_dotenv
from server.helpers.auth import (
    get_user,
    get_password_hash,
    authenticate_user,
    create_csrf_token,
    create_session_id_hash,
    send_forgot_password_email
)
from server.configs.db import users_collection, reset_tokens_collection
from server.helpers.rate_limiter import check_rate_limit
from fastapi.responses import JSONResponse
from Crypto.Cipher import AES
from Crypto.Util.Padding import pad
from jose import jwt, JWTError

load_dotenv()

# oauth2_scheme = OAuth2PasswordBearer(tokenUrl="/auth/login")
router = APIRouter()


@router.get("/health")
async def health_check():
    """Check if the API is running."""
    try:
        return {"status": "ok"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e)) from e


@router.post("/auth/admin/register")
async def register(login_data: LoginInputDataModel):
    """Register a new user with email and password.

    Args:
        login_data (LoginInputDataModel): The user's login data including email and password.

    Returns:
        JSONResponse: A response indicating the registration status.

    Raises:
        HTTPException: If an error occurs during registration.
    """
    try:
        user = await get_user(login_data.email)
        if not user:
            password_hash = get_password_hash(login_data.password)
            credentials = {
                "_id": str(uuid.uuid4()),
                "email": login_data.email,
                "password": password_hash,
                "created_at": datetime.now()
            }
            await users_collection.insert_one(credentials)
            content = {"message": "Registered successfully!!"}
            response = JSONResponse(
                status_code=status.HTTP_201_CREATED, content=content)
            return response

        else:
            response = {"message": "Email address already exists!!"}
            return JSONResponse(status_code=status.HTTP_409_CONFLICT, content=response)

    except Exception as e:
        traceback.print_exc()
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(e)
        ) from e


@router.post("/auth/login")
async def login(loginData: LoginInputDataModel, request: Request):
    try:
        # Check rate limit
        await check_rate_limit(request, "login")

        user = await authenticate_user(loginData.email, loginData.password)
        if not user:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="メールアドレスが見つかりません。",
                headers={"WWW-Authenticate": "Bearer"},
            )
        content = {
            "_id": user["_id"],
            "email": user["email"],
            "message": "ログインしました。"
        }
        CSRF_TOKEN_EXPIRE_DAYS = 90  # 90 days
        CSRF_TOKEN_EXPIRE_MINUTES = CSRF_TOKEN_EXPIRE_DAYS * 24 * 60
        csrf_token_expires = timedelta(minutes=CSRF_TOKEN_EXPIRE_MINUTES)
        session_id = str(uuid.uuid4())
        print("session_id==>", session_id)
        session_id_token = user["email"] + session_id
        csrf_token = create_csrf_token(
            data={

                "email": user["email"],
                "_id": session_id
            },
            secret_key=os.getenv('csrf_token_secrete_key'),
            expires_delta=csrf_token_expires,
        )
        print("csrf_token==>", csrf_token)
        session_hash = create_session_id_hash(session_id_token)
        KEY = bytes(os.getenv('csrf_encryption_secrete_key').encode("utf-8"))
        IV = bytes(os.getenv('aes_encryption_initial_vector').encode("utf-8"))
        cipherText = AES.new(KEY, AES.MODE_CBC, IV)

        cipher_cookie = cipherText.encrypt(
            pad(csrf_token.encode("utf-8"), AES.block_size))
        response = JSONResponse(
            status_code=status.HTTP_200_OK, content=content)

        # Set session cookie
        response.set_cookie(
            key="sessionID",
            value=session_hash,
            max_age=7776000,  # 90 days
            httponly=True,
            samesite="Strict",
            secure=True,
            # domain=".cosbe.inc",
            path="/"
        )

        # Set CSRF token cookie
        response.set_cookie(
            key="__HOST_csrf_token",
            value=cipher_cookie.hex(),
            max_age=7776000,  # 90 days
            samesite="Strict",
            secure=True,
            # domain=".cosbe.inc",
            path="/"
        )

        return response
    except Exception as e:

        traceback.print_exc()
        if e.status_code == 401 or e.status_code == 404:
            raise e
        else:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(
                    e)
            ) from e


@router.post("/auth/forgot-password")
async def forgot_password(forgot_password_data: ForgotPasswordInputDataModel, request: Request):
    """Handle forgot password request and send reset token via email.

    Args:
        forgot_password_data (ForgotPasswordInputDataModel): The user's email.

    Returns:
        JSONResponse: A response indicating the status of the request.
    """
    try:
        # Check rate limit
        await check_rate_limit(request, "forgot_password")

        # Check if the user exists
        user = await get_user(forgot_password_data.email)
        if not user:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="This account is not registered.",
            )

        link_expiration = {"format": "minutes", "value": 30}  # 30 minutes

        # Create a unique token
        print("user==>", user)
        token = create_csrf_token(
            data={
                "_id": user["_id"],
                "email": user["email"],
            },
            secret_key=os.getenv('csrf_token_secrete_key'),
            expires_delta=timedelta(
                **{link_expiration["format"]: link_expiration["value"]}),
        )

        # Construct the reset link with the token
        frontend_url = os.getenv('frontend_url')
        reset_link = f"{frontend_url}/reset-password?token={token}"

        # Send an email to the user with the reset link
        await send_forgot_password_email(user["email"], reset_link, link_expiration)

        content = {
            "message": "パスワードリセットリンクを送信しました。"
        }
        return JSONResponse(status_code=status.HTTP_200_OK, content=content)
    except HTTPException as e:
        raise e
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(e)
        ) from e


@router.post("/auth/reset-password")
async def reset_password(reset_password_data: ResetPasswordInputDataModel, request: Request):
    """Reset the user's password after verifying the token.

    Args:
        reset_password_data (ResetPasswordInputDataModel): The user's new password and token.

    Returns:
        JSONResponse: A response indicating the status of the request.
    """
    try:
        # Check rate limit
        await check_rate_limit(request, "reset_password")

        token = reset_password_data.token
        if not token:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST, detail="Token is missing"
            )

        # Check if token has already been used
        used_token = await reset_tokens_collection.find_one({"token": token})
        if used_token:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="This password reset link has already been used. Please request a new one."
            )

        # Verify the token
        try:
            payload = jwt.decode(
                token, os.getenv('csrf_token_secrete_key'), algorithms=["HS256"])
            email = payload.get("email")
            if email is None:
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST, detail="Invalid token")
        except jwt.ExpiredSignatureError:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST, detail="Token has expired")
        except JWTError:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST, detail="Invalid token")

        # Retrieve the user from the database
        user = await get_user(email)
        if not user:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND, detail="User not found")

        # Update the user with the new password
        password_hash = get_password_hash(reset_password_data.password)
        await users_collection.update_one(
            {"email": email},
            {"$set": {"password": password_hash}}
        )

        # Mark the token as used
        await reset_tokens_collection.insert_one({
            "token": token,
            "email": email,
            "used_at": datetime.now()
        })

        content = {"message": "パスワードをリセットしました。"}
        return JSONResponse(status_code=status.HTTP_200_OK, content=content)

    except HTTPException as e:
        raise e
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(e)
        ) from e


@router.get("/auth/logout")
async def logout(response: Response):

    try:
        response.delete_cookie(
            key="__HOST_csrf_token",
            path="/",
            secure=True,
            # domain=".cosbe.inc",
        )
        response.delete_cookie(
            key="sessionID",
            path="/",
            secure=True,
            # domain=".cosbe.inc",
        )
        return {"message": "ログアウトしました。"}
    except Exception as e:

        traceback.print_exc()
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(e)
        ) from e
