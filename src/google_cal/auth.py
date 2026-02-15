"""Google Calendar OAuth2 authentication and token management."""

import logging
import os
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

logger = logging.getLogger("outlook_gcal_sync.google.auth")

SCOPES = ["https://www.googleapis.com/auth/calendar"]


def get_credentials(credentials_path: str, token_path: str) -> Credentials:
    """Get valid Google OAuth2 credentials.

    Loads from token file if available, refreshes if expired,
    or runs the OAuth flow if no valid credentials exist.

    Args:
        credentials_path: Path to the OAuth client secret JSON.
        token_path: Path to store/load the user's token.

    Returns:
        Valid Google OAuth2 credentials.

    Raises:
        FileNotFoundError: If credentials_path doesn't exist.
    """
    creds = None

    # Load existing token
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
        logger.debug("Loaded existing token from %s", token_path)

    # Refresh or re-authenticate
    if creds and creds.expired and creds.refresh_token:
        logger.info("Refreshing expired Google token...")
        try:
            creds.refresh(Request())
        except Exception as e:
            logger.warning("Token refresh failed: %s. Re-authenticating...", e)
            creds = None

    if not creds or not creds.valid:
        if not os.path.exists(credentials_path):
            raise FileNotFoundError(
                f"Google OAuth credentials file not found: {credentials_path}\n"
                "Please download it from Google Cloud Console and place it at the path above.\n"
                "See: https://console.cloud.google.com/apis/credentials"
            )
        logger.info("Starting Google OAuth flow (browser will open)...")
        flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
        creds = flow.run_local_server(port=0)

    # Save token for next run
    Path(token_path).parent.mkdir(parents=True, exist_ok=True)
    with open(token_path, "w") as token_file:
        token_file.write(creds.to_json())
    logger.debug("Saved token to %s", token_path)

    return creds


def build_calendar_service(credentials_path: str, token_path: str):
    """Build and return an authenticated Google Calendar API service.

    Args:
        credentials_path: Path to the OAuth client secret JSON.
        token_path: Path to store/load the user's token.

    Returns:
        A Google Calendar API service resource.
    """
    creds = get_credentials(credentials_path, token_path)
    service = build("calendar", "v3", credentials=creds)
    logger.info("Google Calendar API service initialized.")
    return service
