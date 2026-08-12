"""
Module: main.py
Description: This is the entry point for the FastAPI application. 
It sets up the API server, configures CORS middleware, and defines a simple health check endpoint. 
"""
import logging

from uuid import uuid4
from time import perf_counter
from collections.abc import Awaitable, Callable
from contextlib import asynccontextmanager

from starlette.exceptions import HTTPException as StarletteHTTPException
from fastapi import FastAPI, Request, Response
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from fastapi.exception_handlers import request_validation_exception_handler
from fastapi.exceptions import RequestValidationError
from src.department import routes as department_routes
from src.team import routes as team_routes
from src.doctor import routes as doctor_routes
from src.db.connection import init_db
from src.utils.logger import configure_logging, request_fields, request_id_var
from src.utils.misc import elapsed_ms

configure_logging()
logger = logging.getLogger(__name__)


@asynccontextmanager
async def lifespan(_app_instance: FastAPI):
    """Handles application startup and shutdown lifecycles safely."""
    logger.info("[Startup] Initializing database tables...")
    init_db()
    yield
    logger.info("[Shutdown] Cleaning up server resources...")


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
app.include_router(doctor_routes.router, prefix="/api/v1")

CallNext = Callable[[Request], Awaitable[Response]]


@app.middleware("http")
async def log_requests(request: Request, call_next):
    """Corellate a request and emit its structured access and error records"""
    request_id = uuid4().hex
    token = request_id_var.set(request_id)
    started = perf_counter()

    try:
        try:
            response = await call_next(request)
        except Exception:
            duration_ms = elapsed_ms(started)
            logger.exception(
                "Unhandled request exception",
                extra=request_fields(
                    request,
                    event="http.error",
                    status_code=500,
                    duration_ms=duration_ms,
                ),
            )

            response = JSONResponse(
                status_code=500,
                content={"detail": "Internal Server Error"}
            )
        duration_ms = elapsed_ms(started)
        logger.info(
            "Request completed",
            extra=request_fields(
                request,
                event="http.request",
                status_code=response.status_code,
                duration_ms=duration_ms,
            ),
        )
        response.headers["X-Request-ID"] = request_id
        return response
    finally:
        request_id_var.reset(token)


@app.exception_handler(StarletteHTTPException)
async def http_exception_handler(request: Request, exc: StarletteHTTPException):
    """Log HTTP exceptions with context before returning the response."""
    fields = request_fields(
        request,
        event="http.error",
        status_code=exc.status_code,
    )
    fields["detail"] = exc.detail

    # Log as WARNING for client errors (4xx), ERROR for server errors (5xx)
    if exc.status_code >= 500:
        logger.error("HTTP %s: %s", exc.status_code, exc.detail, extra=fields)
    else:
        logger.warning("HTTP %s: %s", exc.status_code,
                       exc.detail, extra=fields)

    return JSONResponse(
        status_code=exc.status_code,
        content={"detail": exc.detail},
        headers=getattr(exc, "headers", None),
    )


@app.exception_handler
async def log_validation_error(
    request: Request,
    exc: RequestValidationError
):
    """Log safe validation metadata and preserve FastAPI's error response."""

    # Log the count and route, not exc.body or exc.errors(), because validation
    # input can contain posswords, tokens etc
    fields = request_fields(
        request,
        event="http.validation_error",
        status_code=422,
    )

    fields["error_count"] = len(exc.errors())
    logger.warning("Request validation failed", extra=fields)
    return await request_validation_exception_handler(request, exc)


@app.get("/health", tags=["System"])
async def health_check():
    """Simple baseline route for frontends to check API availability."""
    logger.info("Health check completed", extra={"event": "health.checked"})
    return {"status": "healthy", "engine": "FastAPI + SQLModel"}
