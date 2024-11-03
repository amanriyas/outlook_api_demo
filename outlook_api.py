# Firstly we need to register the app with Azure portal in order to access the Outlook API
# We need to add User.Read under API permissions to get basic details
# OAuth via access tokens
# .env file to store app_id, secret key for the app and tenant_id for organisation

# pip install Flask msal requests python-dotenv

import os
import requests
from flask import Flask, redirect, request, session, jsonify
from dotenv import load_dotenv

load_dotenv() #load environment variables

app = Flask(__name__)
app.secret_key = os.urandom(24) # a secure key for sessions

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
REDIRECT_URI = os.getenv("REDIRECT_URI")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = "User.Read"

@app.route('/')
def home():
    return 'Welcome to the Outlook Sign-In App! <a href="/login">Login with Outlook</a>'

@app.route('/login')
def login():
    # Redirect the user to the Microsoft login page
    auth_url = (
        f"{AUTHORITY}/oauth2/v2.0/authorize?"
        f"client_id={CLIENT_ID}&"
        f"response_type=code&"
        f"redirect_uri={REDIRECT_URI}&"
        f"response_mode=query&"
        f"scope={SCOPE}"
    )
    return redirect(auth_url)

@app.route('/auth')
def auth_callback():
    # Retrieve the authorization code from the callback
    code = request.args.get('code')

    # Exchange the authorization code for an access token
    token_url = f"{AUTHORITY}/oauth2/v2.0/token"
    token_data = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'grant_type': 'authorization_code',
        'code': code,
        'redirect_uri': REDIRECT_URI,
        'scope': SCOPE
    }

    token_response = requests.post(token_url, data=token_data)
    token_response_json = token_response.json()

    # Check for errors in the token response
    if "error" in token_response_json:
        return jsonify(token_response_json), 400

    # Use the access token to get user information from Microsoft Graph API
    access_token = token_response_json['access_token']
    user_info_response = requests.get(
        "https://graph.microsoft.com/v1.0/me",
        headers={"Authorization": f"Bearer {access_token}"}
    )

    # Return the user information as JSON
    return jsonify(user_info_response.json())

if __name__ == '__main__':
    app.run(ssl_context = 'adhoc',debug=True)


