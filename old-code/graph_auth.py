"""
Microsoft Graph authentication (client credentials) for SharePoint/OneDrive access.

Why this exists:
- SharePoint Online uses Azure AD (Entra ID) and requires OAuth2 access tokens.
- For automation, we use an App Registration (service principal) with the client credentials flow.
- This module acquires a token for Microsoft Graph using MSAL.

Environment variables required (set these before running):
- AZURE_TENANT_ID: Your Entra ID tenant GUID (or domain)
- AZURE_CLIENT_ID: The App Registration (Application) ID
- AZURE_CLIENT_SECRET: A client secret for the app registration

Scopes used:
- ["https://graph.microsoft.com/.default"] which maps to the app's granted Graph application permissions.

Usage example:
    from graph_auth import acquire_graph_token
    token = acquire_graph_token()  # reads env vars
    print("Token type:", token.get("token_type"))

Note: Requires the 'msal' package. Install with: pip install msal
"""
from __future__ import annotations
import os
from typing import Dict, Optional

try:
    import msal  # type: ignore
except Exception:
    msal = None  # defer import error to runtime with friendly message

GRAPH_DEFAULT_SCOPE = ["https://graph.microsoft.com/.default"]

class MissingDependencyError(RuntimeError):
    pass

class MissingConfigError(RuntimeError):
    pass


def _require_msal():
    if msal is None:
        raise MissingDependencyError(
            "The 'msal' package is required. Install with: pip install msal"
        )


def acquire_graph_token(
    tenant_id: Optional[str] = None,
    client_id: Optional[str] = None,
    client_secret: Optional[str] = None,
    scopes = GRAPH_DEFAULT_SCOPE,
) -> Dict[str, str]:
    """Acquire an app-only access token for Microsoft Graph.

    Returns the token dict from MSAL (contains 'access_token', 'expires_in', etc.).
    Raises MissingConfigError if env/config is incomplete.
    """
    _require_msal()

    tenant_id = tenant_id or os.getenv("AZURE_TENANT_ID")
    client_id = client_id or os.getenv("AZURE_CLIENT_ID")
    client_secret = client_secret or os.getenv("AZURE_CLIENT_SECRET")

    if not tenant_id or not client_id or not client_secret:
        raise MissingConfigError(
            "AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET must be set."
        )

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=authority,
    )

    result = app.acquire_token_for_client(scopes=scopes)

    # MSAL returns {'error': '...', 'error_description': '...'} on failure
    if "access_token" not in result:
        # include concise details only
        raise RuntimeError(
            f"Failed to acquire token: {result.get('error')}: {result.get('error_description')}"
        )

    return result


def get_bearer_header(token: Dict[str, str]) -> Dict[str, str]:
    """Return Authorization header for requests using the provided token dict."""
    return {"Authorization": f"Bearer {token['access_token']}"}

