import os
import uvicorn
from fastapi import FastAPI
from dotenv import load_dotenv
from mangum import Mangum
from server.api.login import router as login_router
from server.api.projects import router as projects_router
from server.api.tender import router as tender_router
from fastapi.middleware.cors import CORSMiddleware

load_dotenv()

app = FastAPI(title="Coseb Project Management")

# More explicit CORS configuration
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://odashimagumi.cosbe.inc",
                   "http://localhost:5173", "https://odashimagumi.cosbe.inc/"],
    allow_credentials=True, 
    allow_methods=["*"],
    allow_headers=["*"],
    max_age=3600
)

handler = Mangum(app)
app.include_router(login_router, prefix="/api/v1")
app.include_router(projects_router, prefix="/api/v1")
app.include_router(tender_router, prefix="/api/v1/tender")

if __name__ == "__main__":
    port = int(os.getenv('server_port', 8000))  # Default to 8000 if not set
    print(f"The server running on port {port}")
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=port,
        reload=True,
    )
