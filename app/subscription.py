import datetime
import asyncio
from app.email import fetch_email_details
from app.graph_client import get_graph_client
from flask import g

from msgraph.generated.models.subscription import Subscription

async def create_subscription(graph_client, callback_url, lifecycle_url=None):
    """
    Create a subscription to monitor new emails in the inbox using Microsoft Graph SDK.
    :param graph_client: Authenticated Microsoft Graph client.
    :param callback_url: Public URL for webhook notifications.
    :param lifecycle_url: (Optional) URL for lifecycle notifications.
    :return: Subscription details or error message.
    """
    try:
        # Set expiration time (maximum 3 days for delegated permissions)
        expiration_time = (datetime.datetime.utcnow() + datetime.timedelta(days=3)).isoformat() + "Z"

        # Define subscription payload using SDK's Subscription model
        request_body = Subscription(
            change_type="created",
            notification_url=callback_url,
            resource="me/mailFolders('inbox')/messages",
            expiration_date_time=expiration_time,
            client_state="secret-value",
        )

        # Create the subscription
        result = await graph_client.subscriptions.post(request_body)

        # Return the created subscription details
        return {
            "id": result.id,
            "resource": result.resource,
            "expirationDateTime": result.expiration_date_time,
            "notificationUrl": result.notification_url,
        }
    except Exception as e:
        return {"error": str(e)}


async def renew_subscription(graph_client, subscription_id):
    """
    Renew an existing subscription before it expires.
    :param graph_client: Authenticated Microsoft Graph client.
    :param subscription_id: ID of the subscription to renew.
    :return: Renewed subscription details or error message.
    """
    try:
        # Extend the subscription expiration time
        new_expiration_time = (datetime.datetime.utcnow() + datetime.timedelta(days=3)).isoformat() + "Z"

        # Define renewal payload
        renewal_data = {"expirationDateTime": new_expiration_time}

        # Renew subscription
        renewed_subscription = await graph_client.subscriptions[subscription_id].patch(renewal_data)
        return renewed_subscription
    except Exception as e:
        return {"error": str(e)}

def handle_notifications(notification_data):
    """
    Process incoming Microsoft Graph notifications and fetch email details.
    :param notification_data: Notification payload from Microsoft Graph.
    :param graph_client: Authenticated Microsoft Graph client.
    """
    try:
        print(notification_data)
        graph_client = get_graph_client()
        for value in notification_data.get("value", []):
            # Validate clientState for security
            if value.get("clientState") != "secret-value":
                print("Invalid clientState in notification")
                continue

            # Extract message ID
            resource_data = value.get("resourceData", {})
            email_id = resource_data.get("id")
            if not email_id:
                print("No email ID found in notification")
                continue
            # Fetch and process the email asynchronously
            asyncio.run(fetch_email_details(graph_client, email_id))
    except Exception as e:
        print(f"Error handling notification: {str(e)}")
