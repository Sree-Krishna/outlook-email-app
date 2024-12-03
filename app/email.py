from msgraph import GraphServiceClient
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
