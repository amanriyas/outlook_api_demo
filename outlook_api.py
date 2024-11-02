# Firstly we need to register the app with Azure portal in order to access the Outlook API
# We need to add User.Read under API permissions to get basic details
# OAuth via access tokens
# .env file to store app_id, secret key for the app and tenant_id for organisation

# pip install Flask msal requests python-dotenv

import os
from tkinter import Scale

from flask import Flask, request, redirect, url_for, session, jsonify
from msal import ConfidentialClientApplication
from dotenv import load_dotenv
import requests
from sqlalchemy.util.langhelpers import repr_tuple_names

load_dotenv()

app = Flask(__name__)
app.secret_key = os.urandom(24)  #securing the session for each use case

#Credentials for Azure login and Authentication

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
AUTHORITY = os.getenv("AUTHORITY")
REDIRECT_URL = os.getenv("REDIRECT_URI")
SCOPE = os.getenv("SCOPE")

# create an msal instance for confidential client

msal_app = ConfidentialClientApplication(CLIENT_ID,authority=AUTHORITY,client_credential=CLIENT_SECRET)




@app.route("/")
def index():
    """
    Default route and fetches the user data if authenticated via a session token.
    Else it will initiate an Outlook Sign in

    """

    if "token" in session:
        user_data = get_user_profile(session["token"])
        return jsonify(user_data)
    else:
        return '<a href = "\login">"Sign in with Outlook"</a>'

@app.route("/login")
def login():
    """
    generates a authorization url for the user to sign in.
    This will request the user to give basic information
    """
    flow = msal_app.initiate_auth_code_flow(SCOPE,redirect_uri=REDIRECT_URL)
    #Store the flow in the session

    session["flow"] = flow

    return redirect(flow["auth_uri"])

@app.route("/getAToken")
def authorized():
    flow = session.get("flow")
    token_response = msal_app.acquire_token_by_auth_code_flow(flow,request.args)

    # Store the access token in the session

    if "access_token" in token_response:
        session["token"] = token_response["access token"]

    return redirect(url_for("index"))


def get_user_profile(token):
    """
    Fetches user profile information from Microsoft graph api
    The access token is passed in the request header
    This will return the basic user information such as name, username and email.
    Note: Info like Student Id is subject to University restrictions.
    """
    headers = {"Authorization" : "Bearer" + token}
    response = request.get("https://graph.microsoft.com/v1.0/me", headers=headers)
    return response.json()

if __name__ == "__main__":
    app.run(debug=True)

