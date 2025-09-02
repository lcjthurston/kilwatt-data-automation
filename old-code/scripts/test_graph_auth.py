"""
Step 1 test: acquire an app-only Microsoft Graph token using client credentials.

How to run:
1) Ensure you have msal installed:    pip install msal
2) Set environment variables:
   - AZURE_TENANT_ID
   - AZURE_CLIENT_ID
   - AZURE_CLIENT_SECRET
3) Run:  python scripts/test_graph_auth.py
Expected: Prints a short success line with token type and expires_in.
"""
from __future__ import annotations
import json
from graph_auth import acquire_graph_token

def main():
    token = acquire_graph_token()
    print("SUCCESS: acquired Graph token", {
        "token_type": token.get("token_type"),
        "expires_in": token.get("expires_in"),
        "ext_expires_in": token.get("ext_expires_in"),
        "scope": token.get("scope"),
    })

if __name__ == "__main__":
    main()

