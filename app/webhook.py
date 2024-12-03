from flask import Flask, request, jsonify
import logging

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)

@app.route('/webhook', methods=['POST'])
def webhook_listener():
    """
    Webhook endpoint to receive Microsoft Graph notifications.
    """
    data = request.json
    if data:
        logging.info(f"Received notification: {data}")
        # Handle notification processing here
        return jsonify({"status": "success"}), 200
    return jsonify({"error": "Invalid request"}), 400

def start_webhook():
    """
    Start the Flask app to listen for webhook notifications.
    """
    app.run(host="0.0.0.0", port=5000)
