import os
from msal import ConfidentialClientApplication
from config import Config

# Load environment variables
config = Config()
CLIENT_ID = config.CLIENT_ID
CLIENT_SECRET = config.CLIENT_SECRET
TENANT_ID = config.TENANT_ID

# MSAL configuration
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

def get_access_token():
    """
    Acquire an access token using client credentials.
    """
    app = ConfidentialClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    token = app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" in token:
        return token["access_token"]
    raise Exception("Failed to obtain access token")
