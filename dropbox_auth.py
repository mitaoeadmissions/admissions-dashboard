"""
Run this ONCE locally to get your Dropbox refresh token.
After running, save the refresh token as a GitHub Secret.

Usage:  python dropbox_auth.py
"""
import dropbox
from dropbox import DropboxOAuth2FlowNoRedirect

print("=" * 60)
print("  Dropbox One-Time Auth — gets your permanent refresh token")
print("=" * 60)

APP_KEY    = input("\nPaste your Dropbox App Key    : ").strip()
APP_SECRET = input("Paste your Dropbox App Secret : ").strip()

auth_flow = DropboxOAuth2FlowNoRedirect(
    APP_KEY, APP_SECRET, token_access_type='offline'
)
authorize_url = auth_flow.start()

print(f"\n[Step 1] Open this URL in your browser:\n\n  {authorize_url}\n")
print("[Step 2] Click 'Allow'")
print("[Step 3] Copy the code shown on screen\n")

code = input("Paste the authorization code here: ").strip()

try:
    result = auth_flow.finish(code)
    print("\n" + "=" * 60)
    print("  SUCCESS! Save these 3 values as GitHub Secrets:")
    print("=" * 60)
    print(f"\n  DROPBOX_APP_KEY      = {APP_KEY}")
    print(f"  DROPBOX_APP_SECRET   = {APP_SECRET}")
    print(f"  DROPBOX_REFRESH_TOKEN = {result.refresh_token}")
    print("\n" + "=" * 60)
except Exception as e:
    print(f"\nError: {e}")
