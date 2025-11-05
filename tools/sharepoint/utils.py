import requests
import time
import msal
import jwt

LOREM_IPSUM = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."

def get_site(site_url: str, site_path: str, access_token: str) -> str:
    site_endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_url}:{site_path}"
    headers = {'Authorization': f'Bearer {access_token}'}
    return requests.get(site_endpoint, headers=headers).json()['id']

def get_drives(site_id: str, access_token: str):
    drive_endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    headers = {'Authorization': f'Bearer {access_token}'}
    return requests.get(drive_endpoint, headers=headers).json()

def token_expired(token: str) -> bool:
    try:
        exp_timestamp = jwt.decode(token, options={"verify_signature": False}).get("exp", 0)
        return exp_timestamp < time.time()
    except Exception as e:
        print("Error decoding token:", e)
        return True

def device_code_login(tenant_id: str, client_id: str) -> str:
    AUTHORITY = f"https://login.microsoftonline.com/{tenant_id}"
    SCOPES = ["Files.ReadWrite", "Sites.ReadWrite.All"]

    app = msal.PublicClientApplication(
        client_id,
        authority=AUTHORITY
    )

    # Initiates device code flow for user consent
    flow = app.initiate_device_flow(scopes=SCOPES)
    print(flow["message"])
    result = app.acquire_token_by_device_flow(flow)
    return result["access_token"]
