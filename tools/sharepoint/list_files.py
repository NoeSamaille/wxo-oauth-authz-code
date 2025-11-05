import requests

from ibm_watsonx_orchestrate.agent_builder.tools import tool, ToolPermission
from ibm_watsonx_orchestrate.agent_builder.connections import ConnectionType
from ibm_watsonx_orchestrate.run import connections

CONNECTION_MS365="ms365"
CONNECTION_SHAREPOINT="sharepoint"

def list_sharepoint_files_worker(drive_id: str, access_token: str) -> str:
    """
    Function that lists SharePoint files.
    """
    
    url = f'https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children'
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    items = response.json()['value']
    return items

@tool(
    permission=ToolPermission.READ_ONLY,
    expected_credentials=[
        {"app_id": CONNECTION_MS365, "type": ConnectionType.OAUTH2_AUTH_CODE},
        {"app_id": CONNECTION_SHAREPOINT, "type": ConnectionType.KEY_VALUE}
    ],
)
def list_sharepoint_files() -> str:
    """
    Tool list SharePoint files.
    """
    conn = connections.oauth2_auth_code(CONNECTION_MS365)
    access_token = conn.access_token
    
    conn_sharepoint = connections.key_value(CONNECTION_SHAREPOINT)
    site_id = conn_sharepoint.get('site_id')
    drive_id = conn_sharepoint.get('drive_id')

    if not site_id or not drive_id:
        raise ValueError("SharePoint site_id or drive_id not found in credentials.")

    return list_sharepoint_files_worker(drive_id=drive_id, access_token=access_token)

if __name__ == "__main__":
    
    from utils import device_code_login
    from dotenv import load_dotenv
    import os
    
    load_dotenv()
    
    drive_id=os.getenv("SHAREPOINT_DRIVE_ID")
    access_token = device_code_login(os.getenv("ENTRA_TENANT_ID"), os.getenv("ENTRA_CLIENT_ID"))
    print(list_sharepoint_files_worker(drive_id=drive_id, access_token=access_token))
