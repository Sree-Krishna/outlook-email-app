from flask import Flask, request, redirect, g
import asyncio
import datetime
from azure.identity import AuthorizationCodeCredential
from msgraph import GraphServiceClient
from app.subscription import create_subscription, handle_notifications, validate_lifecycle_request
from app.graph_client import get_graph_client
from config import Config

# Flask app for handling OAuth flow and webhooks
app = Flask(__name__)
config = Config()
global graph_client

@app.route("/")
def login():
    """
    Redirect user to Microsoft's authorization endpoint.
    """
    auth_url = f"{config.AUTHORITY_URL}/{config.TENANT_ID}/oauth2/v2.0/authorize"
    params = {
        "client_id": config.CLIENT_ID,
        "response_type": "code",
        "redirect_uri": config.REDIRECT_URI,
        "response_mode": "query",
        "scope": "Mail.Read User.Read",
    }
    query = "&".join([f"{key}={value}" for key, value in params.items()])
    return redirect(f"{auth_url}?{query}")

@app.route("/callback")
def callback():
    """
    Handle the redirect and exchange authorization code for access token.
    Then create a subscription to monitor the user's inbox.
    """
    code = request.args.get("code")
    if not code:
        return "Authorization failed. No code provided."

    try:
        # Initialize Graph client with the authorization code
        graph_client = get_graph_client(auth_code=code)

        # Create a subscription to monitor the user's inbox
        callback_url = config.CALLBACK_URL  # Public URL for notifications
        subscription = asyncio.run(create_subscription(graph_client, callback_url))

        if "error" in subscription:
            return f"Failed to create subscription: {subscription['error']}"
        return f"Subscription created successfully! Subscription ID: {subscription['id']}"
    except Exception as e:
        return f"Failed to authenticate and create subscription: {str(e)}"

@app.route("/notifications", methods=["POST"])
def notifications():
    """
    Handle Microsoft Graph notifications (webhook endpoint).
    """

    if "validationToken" in request.args:
        # Validation token for subscription verification
        return request.args["validationToken"], 200

    # Process the notification payload
    notification_data = request.json
    handle_notifications(notification_data)
    return "", 202

@app.route("/lifecycle", methods=["GET", "POST"])
def lifecycle():
    return validate_lifecycle_request()

if __name__ == "__main__":
    app.run(port=8000, debug=True)
