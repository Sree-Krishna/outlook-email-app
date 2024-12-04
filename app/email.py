from msgraph import GraphServiceClient
from msgraph.generated.users.item.messages.item.reply.reply_post_request_body import ReplyPostRequestBody
from msgraph.generated.users.item.messages.item.reply_all.reply_all_post_request_body import ReplyAllPostRequestBody
from msgraph.generated.models.message import Message
from msgraph.generated.models.recipient import Recipient
from msgraph.generated.models.email_address import EmailAddress
import asyncio

async def fetch_emails(graph_client: GraphServiceClient):
    """Fetch emails from the user's mailbox."""
    try:
        response = await graph_client.me.messages.get()  # Await the async response
        if response:
            html = "<h1>Latest Emails</h1>"
            for message in response.value:
                html += f"<p><strong>Subject:</strong> {message.subject}<br>"
                html += f"<strong>From:</strong> {message.from_.email_address.address}</p><hr>"
            return html
        else:
            return "No emails found or failed to fetch emails."
    except Exception as e:
        return f"Error fetching emails: {str(e)}"

async def fetch_all_emails(graph_client: GraphServiceClient):
    """
    Fetch all emails from the user's inbox with pagination.
    """
    try:
        emails = []
        messages = await graph_client.me.messages.get()
        emails.extend(messages.value)

        # Handle pagination
        while messages.odata_next_link:
            messages = await graph_client.get(messages.odata_next_link)
            emails.extend(messages.value)

        return {"emails": emails}
    except Exception as e:
        return {"error": str(e)}
    
async def fetch_email_details(graph_client: GraphServiceClient, message_id: str):
    """
    Fetch email details (subject, body, sender) using the message ID.
    :param graph_client: Authenticated Microsoft Graph client.
    :param message_id: The ID of the email message to fetch.
    :return: Dictionary containing email details (subject, body, sender).
    """
    try:
        # Fetch the email message using the Graph client

        message = await graph_client.me.messages.by_message_id(message_id).get()

        # Extract email details
        await process_email(message)

    except Exception as e:
        return {"error": str(e)}

async def process_email(message: str):
    """
    Extract and print key details from the email message object.
    """
    try:
        # Extract email metadata
        email_id = message.id
        subject = message.subject
        received_time = message.received_date_time
        sent_time = message.sent_date_time
        from_name = message.from_.email_address.name
        from_address = message.from_.email_address.address

        # Extract recipients
        to_recipients = [
            f"{recipient.email_address.name} <{recipient.email_address.address}>"
            for recipient in message.to_recipients
        ]

        # Extract body content
        body_content = message.body.content
        body_type = message.body.content_type.value  # 'html' or 'text'

        # Email importance
        importance = message.importance.value  # 'normal', 'high', 'low'

        # Web link to email
        web_link = message.web_link

        # Print extracted details
        print(f"Email ID: {email_id}")
        print(f"Subject: {subject}")
        print(f"Received Time: {received_time}")
        print(f"Sent Time: {sent_time}")
        print(f"From: {from_name} <{from_address}>")
        print(f"To: {', '.join(to_recipients)}")
        print(f"Importance: {importance}")
        print(f"Body Type: {body_type}")
        print(f"Body Preview:\n{body_content}")  # Truncated for readability
        print(f"View Email: {web_link}")

    except Exception as e:
        print(f"Error processing email: {str(e)}")

async def reply_message(graph_client: GraphServiceClient, body):
    request_body = ReplyPostRequestBody(
        message = Message(
            to_recipients = [
                Recipient(
                    email_address = EmailAddress(
                        address = body.address,
                        name = body.name,
                    ),
                )
            ],
        ),
        comment = body.comment,
    )

    await graph_client.me.messages.by_message_id('message-id').reply.post(request_body)

async def reply_all_message(graph_client: GraphServiceClient):
    request_body = ReplyAllPostRequestBody(
        comment = "comment-value",
    )

    await graph_client.me.messages.by_message_id('message-id').reply_all.post(request_body)