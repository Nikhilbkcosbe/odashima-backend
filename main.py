import os
import uvicorn
from fastapi import FastAPI
from dotenv import load_dotenv
from mangum import Mangum
from server.api.login import router as login_router
from fastapi.middleware.cors import CORSMiddleware
from server.api.projects import router as projects_router

load_dotenv()

app = FastAPI(title="Coseb Project Management")

# More explicit CORS configuration
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://odajimagumi.cosbe.inc","http://localhost:5173","https://odajimagumi.cosbe.inc/"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    max_age=3600
)

handler = Mangum(app)
app.include_router(login_router, prefix="/api/v1")
app.include_router(projects_router, prefix="/api/v1")   

if __name__ == "__main__":
    print("The server runing with database=", os.getenv('server_port'))
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=int(os.getenv('server_port')),
        reload=True,
    )
