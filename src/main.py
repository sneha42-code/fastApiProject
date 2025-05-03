from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from src.core.config import settings
from src.core.logging import setup_logging
from src.core.cors import setup_cors
from src.api.endpoints import upload, report, health, download

# Initialize logging
logger = setup_logging()

# Create FastAPI app
app = FastAPI(
    title=settings.APP_TITLE,
    description=settings.APP_DESCRIPTION,
    version=settings.APP_VERSION
)

# Setup CORS
setup_cors(app)

# Include routers
app.include_router(health.router, prefix="/api", tags=["health"])
app.include_router(upload.router, prefix="/api", tags=["upload"])
app.include_router(report.router, prefix="/api", tags=["report"])
app.include_router(download.router, prefix="/api", tags=["download"])

# Debug middleware
@app.middleware("http")
async def debug_request(request, call_next):
    logger.info(f"Request: {request.method} {request.url}")
    logger.info(f"Headers: {request.headers}")
    
    response = await call_next(request)
    logger.info(f"Response: {response.status_code}")
    return response

# Root endpoint
@app.get("/")
def read_root():
    return {"message": "Welcome to the Attrition Analysis API"}
