from pydantic import BaseModel, EmailStr


class LoginInputDataModel(BaseModel):
    """Data model for user login input.

    Attributes:
        email (str): The user's email address.
        password (str): The user's password.
    """
    email: EmailStr
    password: str

class ForgotPasswordInputDataModel(BaseModel):
    """Data model for forgot password input.

    Attributes:
        email (str): The user's email address.
    """
    email: EmailStr

class ResetPasswordInputDataModel(BaseModel):
    """Data model for reset password input.

    Attributes:
        password (str): The user's new password.
        token (str): The reset token.
    """
    password: str
    token: str
