"""Microsoft Graph OAuth2 authentication via MSAL interactive browser flow."""

import logging
import os
from pathlib import Path

import msal
import requests as req_lib

logger = logging.getLogger("outlook_gcal_sync.microsoft.auth")

SCOPES = ["Calendars.ReadWrite", "User.Read"]


def _build_http_client() -> req_lib.Session:
    """Build a requests Session that honours corporate CA bundles.

    Checks (in order):
      1. REQUESTS_CA_BUNDLE env var
      2. CURL_CA_BUNDLE env var
      3. SSL_CERT_FILE env var
      4. /tmp/corp_certs.pem (common location for extracted corporate certs)
    Falls back to default verification if none found.
    """
    session = req_lib.Session()

    ca_bundle = (
        os.environ.get("REQUESTS_CA_BUNDLE")
        or os.environ.get("CURL_CA_BUNDLE")
        or os.environ.get("SSL_CERT_FILE")
    )

    # Fallback: check common location for corporate certs
    if not ca_bundle and os.path.exists("/tmp/corp_certs.pem"):
        ca_bundle = "/tmp/corp_certs.pem"

    if ca_bundle and os.path.exists(ca_bundle):
        session.verify = ca_bundle
        logger.info("Using custom CA bundle for MSAL: %s", ca_bundle)
    else:
        logger.debug("No custom CA bundle found, using default SSL verification.")

    return session


def _build_app(client_id: str, tenant_id: str, cache_path: str) -> msal.PublicClientApplication:
    """Build an MSAL PublicClientApplication with persistent token cache."""
    cache = msal.SerializableTokenCache()

    if os.path.exists(cache_path):
        with open(cache_path) as f:
            cache.deserialize(f.read())
        logger.debug("Loaded MSAL token cache from %s", cache_path)

    # Pass custom HTTP client so MSAL uses the corporate CA bundle
    # for its OpenID Connect discovery and token endpoint calls.
    http_client = _build_http_client()

    app = msal.PublicClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        token_cache=cache,
        http_client=http_client,
    )
    return app


def _save_cache(app: msal.PublicClientApplication, cache_path: str) -> None:
    """Persist the token cache to disk if it has changed."""
    cache = app.token_cache
    if cache.has_state_changed:
        Path(cache_path).parent.mkdir(parents=True, exist_ok=True)
        with open(cache_path, "w") as f:
            f.write(cache.serialize())
        logger.debug("Saved MSAL token cache to %s", cache_path)


def get_graph_token(client_id: str, tenant_id: str, cache_path: str) -> str:
    """Get a valid Microsoft Graph access token.

    Attempts silent token acquisition first (cached/refresh token).
    Falls back to interactive browser-based OAuth flow if no valid token exists.
    This uses the authorization code flow (same as Graph Explorer), which works
    even when corporate tenants block device code flow.

    Args:
        client_id: Azure AD application (client) ID.
        tenant_id: Azure AD tenant ID (or "common" for multi-tenant).
        cache_path: Path to the MSAL token cache file.

    Returns:
        A valid access token string.

    Raises:
        RuntimeError: If authentication fails.
    """
    app = _build_app(client_id, tenant_id, cache_path)

    # Try silent acquisition first
    accounts = app.get_accounts()
    if accounts:
        logger.debug("Found %d cached account(s), attempting silent auth...", len(accounts))
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            logger.info("Acquired Microsoft Graph token silently (cached).")
            _save_cache(app, cache_path)
            return result["access_token"]

    # Fall back to interactive browser-based OAuth (like Graph Explorer)
    logger.info("Starting interactive Microsoft browser authentication...")
    print("\nOpening your browser for Microsoft sign-in...")
    print("(If the browser doesn't open automatically, check your taskbar)\n")

    result = app.acquire_token_interactive(scopes=SCOPES)

    if "access_token" not in result:
        error = result.get("error_description", result.get("error", "Unknown error"))
        raise RuntimeError(f"Microsoft authentication failed: {error}")

    _save_cache(app, cache_path)
    logger.info("Microsoft Graph authentication successful.")
    return result["access_token"]
