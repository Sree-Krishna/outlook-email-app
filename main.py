from flask import Flask, request, redirect
from azure.identity import AuthorizationCodeCredential
from msgraph import GraphServiceClient
import asyncio
from app.email import fetch_emails

from config import Config


config = Config()

CLIENT_ID = config.CLIENT_ID
TENANT_ID = config.TENANT_ID
CLIENT_SECRET = config.CLIENT_SECRET
REDIRECT_URI = config.REDIRECT_URI
AUTHORITY_URL = config.AUTHORITY_URL
SCOPES = ["Mail.Read", "User.Read"]

# Flask app for handling OAuth flow
app = Flask(__name__)

@app.route("/")
def login():
    """Redirect user to Microsoft's authorization endpoint."""
    auth_url = f"{AUTHORITY_URL}/oauth2/v2.0/authorize"
    params = {
        "client_id": CLIENT_ID,
        "response_type": "code",
        "redirect_uri": REDIRECT_URI,
        "response_mode": "query",
        "scope": " ".join(SCOPES),
    }
    query = "&".join([f"{key}={value}" for key, value in params.items()])
    return redirect(f"{auth_url}?{query}")

@app.route("/callback")
def callback():
    """Handle the redirect and exchange authorization code for access token."""
    code = request.args.get("code")
    if not code:
        return "Authorization failed. No code provided."

    try:
        # Exchange the authorization code for an access token
        credential = AuthorizationCodeCredential(
            tenant_id=TENANT_ID,
            client_id=CLIENT_ID,
            client_secret=CLIENT_SECRET,
            authorization_code=code,
            redirect_uri=REDIRECT_URI,
        )

        # Initialize the Graph client with the credential
        graph_client = GraphServiceClient(credentials=credential)

        # Use asyncio to call the async function and fetch emails
        emails = asyncio.run(fetch_emails(graph_client))
        return emails
    except Exception as e:
        return f"Failed to authenticate: {str(e)}"

if __name__ == "__main__":
    app.run(port=8000, debug=True)