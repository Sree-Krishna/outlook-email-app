import datetime
import asyncio
from app.email import fetch_email_details
from app.graph_client import get_graph_client
from msgraph import GraphServiceClient
from flask import request
import jsonify

from msgraph.generated.models.subscription import Subscription

async def create_subscription(graph_client: GraphServiceClient, callback_url: str, lifecycle_url=None):
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


async def renew_subscription(graph_client: GraphServiceClient, subscription_id):
    """
    Renew an existing subscription before it expires.
    :param graph_client: Authenticated Microsoft Graph client.
    :param subscription_id: ID of the subscription to renew.
    :return: Renewed subscription details or error message.
    """
    try:
        # Calculate the new expiration time (maximum allowed is 3 days from now)
        new_expiration_time = (datetime.datetime.utcnow() + datetime.timedelta(days=3)).isoformat() + "Z"

        # Create a subscription object with the updated expiration time
        subscription_update = Subscription(expiration_date_time=new_expiration_time)

        # Use the patch method to update the subscription
        renewed_subscription = await graph_client.subscriptions[subscription_id].patch(subscription_update)
        return renewed_subscription
    except Exception as e:
        return {"error": str(e)}

def handle_notifications(notification_data):
    """
    Process incoming Microsoft Graph notifications and fetch email details when a new email is received.
    :param notification_data: Notification payload from Microsoft Graph.
    """
    try:
        print("Received notification:", notification_data)
        graph_client = get_graph_client()

        for value in notification_data.get("value", []):
            # Validate clientState for security
            if value.get("clientState") != "secret-value":
                print("Invalid clientState in notification")
                continue

            # Handle lifecycle events
            if "lifecycleEvent" in value:
                lifecycle_event = value["lifecycleEvent"]
                if lifecycle_event == "missed":
                    print("Missed lifecycle notification. Subscription may be expired.")
                    print(f"Subscription {value['subscriptionId']} expired. Recreating...")
                    callback_url = "<your-callback-url>"  # Replace with your actual callback URL
                    graph_client = get_graph_client()
                    asyncio.run(create_subscription(graph_client, callback_url))
                    continue

            # Check for changeType and process only 'created' notifications
            change_type = value.get("changeType")
            if change_type != "created":
                print(f"Ignoring changeType: {change_type}")
                continue

            # Extract message ID
            resource_data = value.get("resourceData", {})
            email_id = resource_data.get("id")
            if not email_id:
                print("No email ID found in notification")
                continue

            # Fetch and process the email asynchronously
            print(f"Fetching details for new email with ID: {email_id}")
            asyncio.run(fetch_email_details(graph_client, email_id))

    except Exception as e:
        print(f"Error handling notification: {str(e)}")


def validate_lifecycle_request():
    """
    Validates lifecycle requests from Microsoft Graph and responds to lifecycle events.
    """
    try:
        # Microsoft Graph sends a validation token for lifecycle validation
        validation_token = request.args.get("validationToken")
        if validation_token:
            return jsonify(validationToken=validation_token), 200

        # Handle lifecycle events from POST payload
        lifecycle_event = request.json
        if "value" in lifecycle_event:
            for value in lifecycle_event["value"]:
                # Check for lifecycle events
                if "lifecycleEvent" in value:
                    lifecycle_event_type = value["lifecycleEvent"]

                    if lifecycle_event_type == "missed":
                        print("Missed lifecycle notification. Renewing subscription...")
                        subscription_id = value["subscriptionId"]
                        graph_client = get_graph_client()
                        asyncio.run(renew_subscription(graph_client, subscription_id))
                    elif lifecycle_event_type == "deleted":
                        print(f"Subscription {value['subscriptionId']} was deleted. Recreating...")
                        callback_url = "<your-callback-url>"  # Replace with your actual callback URL
                        graph_client = get_graph_client()
                        asyncio.run(create_subscription(graph_client, callback_url))
                        continue
        return "", 202

    except Exception as e:
        return {"error": str(e)}, 500