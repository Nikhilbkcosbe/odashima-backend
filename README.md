# Backend Odajimagumi

A FastAPI-based backend service with authentication and user management features.

## Features

- User Authentication
  - Login with email and password
  - Admin registration
  - Password reset functionality
  - Secure session management with CSRF protection
  - Rate limiting for API endpoints
- Security Features
  - Password hashing
  - CSRF token protection
  - Secure cookie handling
  - Rate limiting
  - JWT token-based authentication

## API Endpoints

### Authentication
- `POST /auth/admin/register` - Register a new admin user
- `POST /auth/login` - User login
- `POST /auth/forgot-password` - Request password reset
- `POST /auth/reset-password` - Reset password with token
- `GET /auth/logout` - User logout
- `GET /health` - Health check endpoint

## Prerequisites

- Python 3.8+
- MongoDB
- SMTP server for email functionality

## Environment Variables

Create a `.env` file in the root directory with the following variables:

```env
# Database
MONGODB_URI=your_mongodb_connection_string

# Security
csrf_token_secrete_key=your_csrf_secret_key
csrf_encryption_secrete_key=your_encryption_key
aes_encryption_initial_vector=your_aes_iv

# Frontend
frontend_url=your_frontend_url

# Email
SMTP_HOST=your_smtp_host
SMTP_PORT=your_smtp_port
SMTP_USER=your_smtp_username
SMTP_PASSWORD=your_smtp_password
```

## Installation

1. Clone the repository:
```bash
git clone https://github.com/your-username/backend-odajimagumi.git
cd backend-odajimagumi
```

2. Create and activate a virtual environment:
```bash
python -m venv venv
# On Windows
venv\Scripts\activate
# On Unix or MacOS
source venv/bin/activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Set up environment variables:
- Copy `.env.example` to `.env`
- Fill in the required environment variables

5. Run the application:
```bash
uvicorn server.main:app --reload
```

The server will start at `http://localhost:8000`

## API Documentation

Once the server is running, you can access:
- Swagger UI documentation: `http://localhost:8000/docs`
- ReDoc documentation: `http://localhost:8000/redoc`

## Security Features

1. **Rate Limiting**
   - Login attempts: 5 requests per minute
   - Password reset: 5 requests per minute
   - Other endpoints: Configurable limits

2. **CSRF Protection**
   - CSRF tokens are required for sensitive operations
   - Tokens expire after 90 days
   - Secure cookie settings

3. **Password Security**
   - Passwords are hashed before storage
   - Password reset tokens expire after 30 minutes
   - One-time use password reset tokens

## Development

### Project Structure
```
backend/
├── server/
│   ├── api/
│   │   └── login.py
│   ├── configs/
│   │   └── db.py
│   ├── helpers/
│   │   ├── auth.py
│   │   └── rate_limiter.py
│   ├── schemas/
│   │   └── login.py
│   └── main.py
├── .env
├── requirements.txt
└── README.md
```

### Running Tests
```bash
pytest
```

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.
