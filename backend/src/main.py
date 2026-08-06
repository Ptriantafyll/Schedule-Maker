"""
Module: main.py
Description: This is the entry point for the FastAPI application. 
It sets up the API server, configures CORS middleware, and defines a simple health check endpoint. 
"""

from contextlib import asynccontextmanager
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from src.department import routes as department_routes
from src.team import routes as team_routes
from src.doctor import routes as doctor_routes
from src.db.connection import init_db


@asynccontextmanager
async def lifespan(_app_instance: FastAPI):
    """Handles application startup and shutdown lifecycles safely."""
    print("[Startup] Initializing database tables...")
    init_db()
    yield
    print("[Shutdown] Cleaning up server resources...")


app = FastAPI(
    title="Hospital Shift Scheduler API",
    description="Backend optimization engine and data sync portal for scheduling duties.",
    version="1.0.0",
    lifespan=lifespan,
)

# Configure CORS so Flutter Web and Mobile can reach this API
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(department_routes.router, prefix="/api/v1")
app.include_router(team_routes.router, prefix="/api/v1")
app.include_router(doctor_routes.router, prefix="/api/v1" )

@app.get("/health", tags=["System"])
async def health_check():
    """Simple baseline route for frontends to check API availability."""
    return {"status": "healthy", "engine": "FastAPI + SQLModel"}
