import os
import traceback
from fastapi import HTTPException, status
from fastapi_mail import FastMail, MessageSchema
from server.constants.auth import conf
from jinja2 import Template
import smtplib
from datetime import datetime


async def send_email(recipient_email, subject, body, body_type):
    """Send an email to the recipient.

    Args:
        recipient_email (str): The email address of the recipient.
        subject (str): The subject of the email.
    """

    try:
        fm = FastMail(conf)
        await fm.send_message(
            MessageSchema(
                subject=subject,
                recipients=recipient_email,
                body=body,
                subtype=body_type,
            )
        )

    except (smtplib.SMTPException, smtplib.SMTPRecipientsRefused) as e:
        # Handle email-related exceptions
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Failed to send email: {str(e)}"
        )
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(e)
        ) from e


async def send_invitation_email(recipient_email, setup_link, link_expiration):
    """Send an invitation email to the recipient to set up their account.

    Args:
        recipient_email (str): The email address of the recipient.
        setup_link (str): The link to set up the account password.
    """
    try:
        # Get the current working directory
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # Construct the path to the email template file
        template_path = os.path.join(
            current_dir, "../templates/invitation_email.html")

        with open(template_path, "r", encoding="utf-8") as file:
            template = Template(file.read())

        body = template.render(
            setup_link=setup_link, link_expiration_format=link_expiration["format"], link_expiration_value=link_expiration['value'])
        subject = "アカウント作成のご案内"
        await send_email([recipient_email], subject, body, "html")

    except Exception as e:
        traceback.print_exc()
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(e)
        ) from e


async def send_task_creation_email(recipient_email, task_data):
    """Send a task creation notification email to the recipient.

    Args:
        recipient_email (str): The email address of the recipient.
        task_data (dict): The task data containing task details.
    """
    try:
        # Get the current working directory
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # Construct the path to the email template file
        template_path = os.path.join(
            current_dir, "../templates/task_creation_email.html")

        with open(template_path, "r", encoding="utf-8") as file:
            template = Template(file.read())

        # Format dates for display
        start_date = task_data["start"].strftime(
            "%Y-%m-%d") if isinstance(task_data["start"], datetime) else task_data["start"]
        end_date = task_data["end"].strftime(
            "%Y-%m-%d") if isinstance(task_data["end"], datetime) else task_data["end"]

        # Render the template with task data
        body = template.render(
            task_name=task_data["text"],
            task_description=task_data["task_description"],
            start_date=start_date,
            end_date=end_date,
            assignee=task_data["assignee"],
            progress=task_data["progress"],
            # Replace with your actual task URL
            task_link=os.getenv("frontend_url")
        )

        subject = "新規タスク作成のお知らせ"
        await send_email([recipient_email], subject, body, "html")

    except Exception as e:
        traceback.print_exc()
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(e)
        ) from e


async def send_assignee_change_email(recipient_email, task_data, old_assignee, new_assignee):
    """Send an assignee change notification email to the recipient.

    Args:
        recipient_email (str): The email address of the recipient.
        task_data (dict): The task data containing task details.
        old_assignee (str): The previous assignee's email.
        new_assignee (str): The new assignee's email.
    """
    try:
        # Get the current working directory
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # Construct the path to the email template file
        template_path = os.path.join(
            current_dir, "../templates/assignee_change_email.html")

        with open(template_path, "r", encoding="utf-8") as file:
            template = Template(file.read())

        # Format dates for display
        start_date = task_data["start"].strftime(
            "%Y-%m-%d") if isinstance(task_data["start"], datetime) else task_data["start"]
        end_date = task_data["end"].strftime(
            "%Y-%m-%d") if isinstance(task_data["end"], datetime) else task_data["end"]

        # Render the template with task data
        body = template.render(
            task_name=task_data["text"],
            task_description=task_data["task_description"],
            start_date=start_date,
            end_date=end_date,
            progress=task_data["progress"],
            old_assignee=old_assignee,
            new_assignee=new_assignee,
            # Replace with your actual task URL
            task_link=os.getenv("frontend_url")
        )

        subject = "タスク担当者変更のお知らせ"
        await send_email([recipient_email], subject, body, "html")

    except Exception as e:
        traceback.print_exc()
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(e)
        ) from e


async def send_task_start_email(recipient_email, task_data):
    """Send a task start notification email to the recipient.

    Args:
        recipient_email (str): The email address of the recipient.
        task_data (dict): The task data containing task details.
    """
    try:
        # Get the current working directory
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # Construct the path to the email template file
        template_path = os.path.join(
            current_dir, "../templates/task_start_email.html")

        with open(template_path, "r", encoding="utf-8") as file:
            template = Template(file.read())

        # Format dates for display
        start_date = task_data["start"].strftime(
            "%Y-%m-%d") if isinstance(task_data["start"], datetime) else task_data["start"]
        end_date = task_data["end"].strftime(
            "%Y-%m-%d") if isinstance(task_data["end"], datetime) else task_data["end"]

        # Render the template with task data
        body = template.render(
            task_name=task_data["text"],
            task_description=task_data["task_description"],
            start_date=start_date,
            end_date=end_date,
            assignee=task_data["assignee"],
            progress=task_data["progress"],
            task_link=os.getenv("frontend_url")
        )

        subject = "タスク開始のお知らせ"
        await send_email([recipient_email], subject, body, "html")

    except Exception as e:
        traceback.print_exc()
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(e)
        ) from e


async def send_task_completion_email(recipient_email, task_data, is_next_task=False):
    """Send a task completion notification email to the recipient.

    Args:
        recipient_email (str): The email address of the recipient.
        task_data (dict): The task data containing task details.
        is_next_task (bool): Whether this is a notification for the next task.
    """
    try:
        # Get the current working directory
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # Construct the path to the email template file
        template_path = os.path.join(
            current_dir, "../templates/task_completion_email.html")

        with open(template_path, "r", encoding="utf-8") as file:
            template = Template(file.read())

        # Format dates for display
        start_date = task_data["start"].strftime(
            "%Y-%m-%d") if isinstance(task_data["start"], datetime) else task_data["start"]
        end_date = task_data["end"].strftime(
            "%Y-%m-%d") if isinstance(task_data["end"], datetime) else task_data["end"]

        # Render the template with task data
        body = template.render(
            task_name=task_data["text"],
            task_description=task_data["task_description"],
            start_date=start_date,
            end_date=end_date,
            assignee=task_data["assignee"],
            progress=task_data["progress"],
            is_next_task=is_next_task,
            task_link=os.getenv("frontend_url")
        )

        subject = "タスク完了のお知らせ"
        await send_email([recipient_email], subject, body, "html")

    except Exception as e:
        traceback.print_exc()
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(e)
        ) from e
