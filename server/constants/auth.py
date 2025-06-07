import os
from dotenv import load_dotenv
from fastapi_mail import ConnectionConfig

load_dotenv()

ORIGINS = [
    'https://project-management.cosbe.inc',
    'https://www.project-management.cosbe.inc',
    'http://localhost:3000', 
    'http://localhost:5173',
    'http://localhost:8000'
    ]

REFERRERS = [
    'http://localhost',
    'https://project-management.cosbe.inc/',
    'https://www.project-management.cosbe.inc/',
    'http://localhost:3000/',
    'http://localhost:5173/'
    ]

conf = ConnectionConfig(
    MAIL_USERNAME=os.getenv("mail_username"),
    MAIL_PASSWORD=os.getenv("mail_password"),
    MAIL_PORT=465,
    MAIL_SERVER=os.getenv("mail_server"),
    MAIL_STARTTLS=False,
    MAIL_SSL_TLS=True,
    USE_CREDENTIALS=True,
    MAIL_FROM=os.getenv("mail_from")
)