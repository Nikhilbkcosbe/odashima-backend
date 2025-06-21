import os
from typing import Dict
from datetime import datetime, timedelta
from typing import Optional
from server.constants.auth import ORIGINS, REFERRERS
from fastapi import HTTPException, Request, status, Depends
from fastapi.security import OAuth2, OAuth2PasswordBearer
from passlib.context import CryptContext
from jose import jwt, JWTError
from fastapi.openapi.models import OAuthFlows as OAuthFlowsModel
from Crypto.Cipher import AES
from dotenv import load_dotenv
from server.configs.db import users_collection
from jinja2 import Template
from fastapi_mail import FastMail, MessageSchema
from server.constants.auth import conf

load_dotenv()

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 30


def verify_password(plain_password, hashed_password):
    """Verify a plain password against a hashed password."""
    return pwd_context.verify(plain_password, hashed_password)


def get_password_hash(password):
    """Hash a password using bcrypt."""
    return pwd_context.hash(password)


async def get_user(email: str):
    """Retrieve a user by email from the database.

    Args:
        email (str): The email of the user to retrieve.

    Returns:
        dict: A dictionary containing user data if found.

    Raises:
        HTTPException: If the user does not have a password or email.
    """
    user = await users_collection.find_one({"email": email})
    if user:
        if "password" in user and "email" in user:
            data = {
                "_id": user["_id"],
                "email": user["email"],
                "hashed_password": user["password"],
            }
            return data
        else:
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Password is not created for this account still. Please complete account creation.",
                headers={"WWW-Authenticate": "Bearer"},
            )


async def authenticate_user(email: str, password: str):
    user = await get_user(email)
    if not user:
        return False
    if not verify_password(password, user["hashed_password"]):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="パスワードが間違っています。",
            headers={"WWW-Authenticate": "Bearer"},
        )
    return user


def create_csrf_token(data: dict, secret_key: str, expires_delta: Optional[timedelta] = None):
    to_encode = data.copy()
    if expires_delta:
        expire = datetime.utcnow() + expires_delta
    else:
        expire = datetime.utcnow() + timedelta(minutes=15)
    to_encode["exp"] = expire
    return jwt.encode(to_encode, secret_key, algorithm=ALGORITHM)


def create_session_id_hash(csrf_token):
    return pwd_context.hash(csrf_token)


def unpad(s): return s[:-ord(s[len(s) - 1:])]


class OAuth2PasswordBearerWithCookie(OAuth2):
    def __init__(
            self,
            tokenUrl: str,
            scheme_name: Optional[str] = None,
            scopes: Optional[Dict[str, str]] = None,
            auto_error: bool = True,
    ):
        if not scopes:
            scopes = {}
        flows = OAuthFlowsModel(
            password={"tokenUrl": tokenUrl, "scopes": scopes})
        super().__init__(flows=flows, scheme_name=scheme_name, auto_error=auto_error)

    async def __call__(self, request: Request) -> Optional[str]:

        origins = ORIGINS
        referers = REFERRERS

        KEY = bytes(os.getenv('csrf_encryption_secrete_key').encode('utf-8'))
        IV = bytes(os.getenv('aes_encryption_initial_vector').encode('utf-8'))
        try:

            csrf_cookie = request.cookies.get("__HOST_csrf_token")
            session_cookie = request.cookies.get("sessionID")
            csrf_header_token = request.headers["X-CSRF-TOKEN"]

            if request.headers["origin"] in origins:

                try:
                    # Security Level 1(If any of the below argument not found in respective place, then it's
                    # unauthorized)
                    if csrf_cookie and bool(csrf_header_token) and session_cookie:
                        cipher = AES.new(KEY, AES.MODE_CBC, IV)
                        try:
                            # Security Level 2(If decryption fail because of cookie tamper, then it's unauthorized)
                            plainText = unpad(cipher.decrypt(
                                bytes.fromhex(csrf_cookie)))

                            try:
                                # Security Level 3(If JWT token failed decode inside the decrypted text,then it's
                                # unauthorized)
                                decoded_subjects = jwt.decode(
                                    plainText, os.getenv('csrf_token_secrete_key'), algorithms=["HS256"])
                                # Security Level 4(If session ID inside the JWT sub of decrypted text !=session ID in
                                # request header,then it's unauthorized)
                                if decoded_subjects["_id"] == csrf_header_token:
                                    print(
                                        "CSRF verification Done!!.Session match waiting.....")
                                    hash_sub = decoded_subjects["email"] + \
                                        decoded_subjects["_id"]
                                    try:
                                        # Security Level 5(if token is not tied to respective session(i.e use of
                                        # someone cookie in someone's browser),then it's unauthorized)
                                        if pwd_context.verify(hash_sub, session_cookie):
                                            print(
                                                "CSRF verified and session matched")
                                            return decoded_subjects
                                        else:
                                            raise HTTPException(
                                                status_code=status.HTTP_401_UNAUTHORIZED,
                                                detail="Session Match Failed",
                                                headers={
                                                    "WWW-Authenticate": "Bearer"},
                                            )
                                    except Exception as e:
                                        print(e)
                                        raise HTTPException(
                                            status_code=status.HTTP_401_UNAUTHORIZED,
                                            detail="Session Match Failed",
                                            headers={
                                                "WWW-Authenticate": "Bearer"},
                                        )

                                else:
                                    raise HTTPException(
                                        status_code=status.HTTP_401_UNAUTHORIZED,
                                        detail="CSRF verify failed!!",
                                        headers={"WWW-Authenticate": "Bearer"},
                                    )

                            except Exception as e:
                                print(e)
                                raise HTTPException(
                                    status_code=status.HTTP_401_UNAUTHORIZED,
                                    detail="Not authenticated",
                                    headers={"WWW-Authenticate": "Bearer"},
                                )
                        except Exception as e:
                            print(e)
                            raise HTTPException(
                                status_code=status.HTTP_401_UNAUTHORIZED,
                                detail="Not authenticated",
                                headers={"WWW-Authenticate": "Bearer"},
                            )
                except Exception as e:
                    print(e)
                    raise HTTPException(
                        status_code=status.HTTP_401_UNAUTHORIZED,
                        detail="All auth parmater not found",
                        headers={"WWW-Authenticate": "Bearer"},
                    )

        except KeyError as e:

            if e:

                if request.headers["referer"] in referers:

                    try:
                        cipher = AES.new(KEY, AES.MODE_CBC, IV)
                        plainText = unpad(cipher.decrypt(
                            bytes.fromhex(csrf_cookie)))
                        decoded_subjects = jwt.decode(
                            plainText, os.getenv('csrf_token_secrete_key'), algorithms=["HS256"])

                        try:
                            hash_sub = decoded_subjects["email"] + \
                                decoded_subjects["_id"]

                            if pwd_context.verify(hash_sub, session_cookie):
                                print("CSRF verified and session matched")
                                return decoded_subjects
                            else:
                                raise HTTPException(
                                    status_code=status.HTTP_401_UNAUTHORIZED,
                                    detail="Session Match Failed",
                                    headers={"WWW-Authenticate": "Bearer"}, )
                        except Exception as e:
                            print(e)
                            raise HTTPException(
                                status_code=status.HTTP_401_UNAUTHORIZED,
                                detail="Session Match Failed",
                                headers={"WWW-Authenticate": "Bearer"}, )
                    except Exception as e:
                        print(e)
                        raise HTTPException(
                            status_code=status.HTTP_401_UNAUTHORIZED,
                            detail="Session Match Failed",
                            headers={"WWW-Authenticate": "Bearer"},
                        )



async def send_forgot_password_email(recipient_email: str, reset_link: str, link_expiration: dict):
    """Send a forgot password email to the specified email address.

    Args:
        recipient_email (str): The recipient's email address.
        reset_link (str): The password reset link.
        link_expiration (dict): Dictionary containing link expiration details.
    """
    try:
        # Get the current working directory
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # Construct the path to the email template file
        template_path = os.path.join(
            current_dir, "../templates/forgot_password.html")

        with open(template_path, "r", encoding="utf-8") as file:
            template = Template(file.read())

        # Prepare template context
        context = {
            "reset_link": reset_link,
            "link_expiration": link_expiration
        }

        # Render the template
        body = template.render(**context)

        # Set subject
        subject = "パスワードリセットのお知らせ"

        # Send email using FastMail with proper encoding
        fm = FastMail(conf)
        await fm.send_message(
            MessageSchema(
                subject=subject,
                recipients=[recipient_email],
                body=body,
                subtype="html",
                charset="utf-8"
            )
        )

    except Exception as e:
        print(f"Error sending forgot password email: {str(e)}")
        raise e from e
