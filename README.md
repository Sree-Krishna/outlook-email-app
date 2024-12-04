
# Microsoft Graph Email Notification Application

This application listens for new emails in a Microsoft Outlook account using the Microsoft Graph API. It processes notifications about new emails and retrieves email details, such as the subject, sender, and body content. Built using Flask and the Microsoft Graph Python SDK, the application demonstrates secure authentication, subscription management, and webhook handling.

## Features

- Secure OAuth 2.0 authentication using the authorization code flow.
- Subscribes to notifications for new emails using the Microsoft Graph API.
- Automatically validates and handles webhook requests.
- Fetches email details, including sender, subject, and body content, upon notification.
- Modular and scalable code structure for deployment in Azure or other cloud platforms.

---

## Prerequisites

1. **Microsoft Azure Account**
   - Register an application in Azure Active Directory to use the Microsoft Graph API.
2. **Microsoft Graph API Permissions**
   - Delegated permissions: `Mail.Read` and `User.Read`.
3. **Environment**
   - Python 3.8+ installed.
   - Publicly accessible endpoint (e.g., [ngrok](https://ngrok.com/) or Azure deployment).

---

## Installation

### Clone the Repository

```bash
git clone https://github.com/your-repo-name/graph-email-notification.git
cd graph-email-notification
```

### Create a Virtual Environment

```bash
python -m venv venv
source venv/bin/activate  # Linux/macOS
venv\Scripts\activate     # Windows
```

### Install Dependencies

```bash
pip install -r requirements.txt
```

### Configure Environment Variables

1. Create a `.env` file in the project root:

```env
CLIENT_ID=<your-client-id>
TENANT_ID=<your-tenant-id>
CLIENT_SECRET=<your-client-secret>
REDIRECT_URI=http://localhost:8000/callback
CALLBACK_URL=<your-public-notification-url>/notifications
LIFECYCLE_URL=<your-public-lifecycle-url>/lifecycle
```

2. Replace the placeholders (`<your-client-id>`, etc.) with the values from your Azure application registration.

---

## Running the Application

### Start the Flask Server

```bash
python main.py
```

### Expose the Application Publicly (Optional for Local Testing)

Use [ngrok](https://ngrok.com/) to expose your local server for webhook callbacks:

```bash
ngrok http 8000
```

Update your `CALLBACK_URL` and `LIFECYCLE_URL` in the `.env` file with the generated ngrok URL.

---

## Usage

1. Open your browser and navigate to `http://localhost:8000/`.
2. Log in with your Microsoft account to authenticate and authorize the application.
3. A subscription will be created to listen for new emails.
4. On receiving a new email, the application processes the notification and retrieves email details.

---

## Project Structure

```
graph-email-notification/
├── app/
│   ├── __init__.py        # Initializes the Flask app
│   ├── graph_client.py    # Handles Microsoft Graph client authentication
│   ├── subscription.py    # Manages subscription creation and renewal
│   ├── email.py           # Fetches email details from Microsoft Graph
│   └── utils.py           # Utility functions for logging and error handling
├── config.py              # Configuration settings and environment variable loading
├── main.py                # Main Flask application
├── requirements.txt       # Python dependencies
└── README.md              # Project documentation
```

---

## Endpoints

### `/`
Redirects users to Microsoft's OAuth 2.0 authorization endpoint.

### `/callback`
Handles the OAuth 2.0 redirect, exchanges the authorization code for tokens, and creates a subscription.

### `/notifications`
Webhook endpoint to handle new email notifications from Microsoft Graph.

### `/lifecycle`
Optional endpoint to handle lifecycle events such as subscription expiration.

---

## Deployment

### Deploy to Azure App Service

1. Login to Azure CLI:
   ```bash
   az login
   ```

2. Deploy the application:
   ```bash
   az webapp up --name <your-app-name> --resource-group <your-resource-group> --runtime "PYTHON:3.9"
   ```

3. Update the `CALLBACK_URL` and `LIFECYCLE_URL` in the `.env` file with the Azure App Service URL.

---

## Future Enhancements

- Add support for managing attachments in emails.
- Implement retry logic for failed notifications.
- Enhance the UI with a dashboard for email and subscription management.
- Automate subscription renewal before expiration.

---

## Troubleshooting

### Common Issues

1. **Subscription Validation Fails**
   - Ensure your `CALLBACK_URL` is publicly accessible and responds with the `validationToken` when requested.

2. **Token Expired Errors**
   - Check if the `AuthorizationCodeCredential` is correctly managing token refresh.

3. **Unable to Receive Notifications**
   - Verify that the subscription is active and webhook endpoint is reachable.

---

## License

This project is licensed under the MIT License. See the LICENSE file for details.

---

## Acknowledgments

- [Microsoft Graph Python SDK](https://github.com/microsoftgraph/msgraph-sdk-python)
- [Flask](https://flask.palletsprojects.com/)
- [Azure Identity](https://pypi.org/project/azure-identity/)
