from fastapi.middleware.cors import CORSMiddleware

def setup_cors(app):
    """
    Set up CORS middleware for the FastAPI application.
    """
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["http://localhost:3000", "http://automateroperting", 
                      "https://sneha42-code.github.io", "*"],
        allow_credentials=True,
        allow_methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"],
        allow_headers=["*"],
        expose_headers=["*"]
    )