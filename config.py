from dotenv import load_dotenv
import os

# Load environment variables from .env file
load_dotenv()

class Config:
    CLIENT_ID = os.getenv("CLIENT_ID")
    CLIENT_SECRET = os.getenv("CLIENT_SECRET")
    TENANT_ID = os.getenv("TENANT_ID")
    WEBHOOK_URL = os.getenv("WEBHOOK_URL")
    REDIRECT_URI = os.getenv("REDIRECT_URI")

    # Additional configurations can go here
    GRAPH_SCOPES = ["https://graph.microsoft.com/.default"]
    GRAPH_API_BASE_URL = "https://graph.microsoft.com/v1.0"
    AUTHORITY_URL = f"https://login.microsoftonline.com/{TENANT_ID}"
