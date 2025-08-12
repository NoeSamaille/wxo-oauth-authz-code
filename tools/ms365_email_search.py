import requests
from typing import Optional
from pydantic import Field, BaseModel

from ibm_watsonx_orchestrate.agent_builder.tools import tool, ToolPermission
from ibm_watsonx_orchestrate.run import connections
from ibm_watsonx_orchestrate.agent_builder.connections import ConnectionType


CONNECTION_MS365 = "ms365"


@tool(
    name="ms365_email_search",
    description="""
        Search user emails from Microsoft Outlook folder
        
        Example search properties format:
        
        Property,Examples
        AttachmentNames,attachmentnames:annualreport.ppt OR attachmentnames:ppt OR attachmentnames:annual*
        Bcc,bcc:pilarp@contoso.com OR bcc:pilarp OR bcc:Pilar Pinilla
        Category,category:Red Category
        Cc,cc:pilarp@contoso.com OR cc:pilarp OR cc:Pilar Pinilla
        Folderid,folderid:4D6DD7F943C29041A65787E30F02AD1F00000000013A0000 OR folderid:4D6DD7F943C29041A65787E30F02AD1F00000000013A0000 AND from:garthf@contoso.com
        From,from:pilarp@contoso.com OR from:pilarp OR from:Pilar Pinilla
        HasAttachment,from:pilar@contoso.com AND hasAttachment:true OR hasAttachment:false
        Importance,importance:high OR importance:medium OR importance:low
        IsRead,isread:true OR isread:false
        ItemClass,itemclass:ipm.externaldata.Facebook* AND subject:contoso OR itemclass:ipm.externaldata.Twitter* AND from:Ann Beebe AND Northwind Traders
        Kind,kind:email OR kind:email OR kind:im OR kind:externaldata
        Participants,participants:garthf@contoso.com OR participants:contoso.com
        Received,received:2021-04-15 OR received>=2021-01-01 AND received<=2021-03-31
        Recipients,recipients:garthf@contoso.com OR recipients:contoso.com
        Sent,sent:2021-07-01 OR sent>=2021-06-01 AND sent<=2021-07-01
        Size,size>26214400 OR size:1..1048567
        Subject,subject:Quarterly Financials OR subject:northwind
        To,to:annb@contoso.com OR to:annb OR to:Ann Beebe
        
    """,
    permission=ToolPermission.READ_ONLY,
    expected_credentials=[
        {"app_id": CONNECTION_MS365, "type": ConnectionType.OAUTH2_AUTH_CODE}
    ],
)
def email_search(
    search: Optional[str] = '',
    email_folder: Optional[str] = "inbox",
    top: int = 1
) -> list:
    """
    Get Microsoft 365 calendar events.
        
    Example search properties format:
    
    Property,Examples
    AttachmentNames,attachmentnames:annualreport.ppt OR attachmentnames:ppt OR attachmentnames:annual*
    Bcc,bcc:pilarp@contoso.com OR bcc:pilarp OR bcc:Pilar Pinilla
    Category,category:Red Category
    Cc,cc:pilarp@contoso.com OR cc:pilarp OR cc:Pilar Pinilla
    Folderid,folderid:4D6DD7F943C29041A65787E30F02AD1F00000000013A0000 OR folderid:4D6DD7F943C29041A65787E30F02AD1F00000000013A0000 AND from:garthf@contoso.com
    From,from:pilarp@contoso.com OR from:pilarp OR from:Pilar Pinilla
    HasAttachment,from:pilar@contoso.com AND hasAttachment:true OR hasAttachment:false
    Importance,importance:high OR importance:medium OR importance:low
    IsRead,isread:true OR isread:false
    ItemClass,itemclass:ipm.externaldata.Facebook* AND subject:contoso OR itemclass:ipm.externaldata.Twitter* AND from:Ann Beebe AND Northwind Traders
    Kind,kind:email OR kind:email OR kind:im OR kind:externaldata
    Participants,participants:garthf@contoso.com OR participants:contoso.com
    Received,received:2021-04-15 OR received>=2021-01-01 AND received<=2021-03-31
    Recipients,recipients:garthf@contoso.com OR recipients:contoso.com
    Sent,sent:2021-07-01 OR sent>=2021-06-01 AND sent<=2021-07-01
    Size,size>26214400 OR size:1..1048567
    Subject,subject:Quarterly Financials OR subject:northwind
    To,to:annb@contoso.com OR to:annb OR to:Ann Beebe
    
    Args:
        search (string): Microsft 365 formatted search query.
        email_folder (string): Outlook email folder to search from. Default: 'inbox'
        top (int): The number of emails to search. Default: 1

    Returns:
        list: A list of calendar events.
    """
    conn = connections.oauth2_auth_code(CONNECTION_MS365)
    
    # Build email properties
    params = {
        '$select': 'subject,from,sender,toRecipients,receivedDateTime,isRead,hasAttachments,importance,body',
        '$top': top,
        '$search': f'"{search}"',
        'mailFolderId': email_folder
    }
    
    # Graph API endpoint for calendar events in the specified range
    url = f"https://graph.microsoft.com/v1.0/me/mailFolders('{email_folder}')/messages"

    # Set headers with the OAuth2 access token
    headers = {
        "Authorization": f"Bearer {conn.access_token}",
        "Prefer": 'outlook.timezone="UTC"',
    }

    # Make the API request
    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()
    emails = response.json().get("value", [])
    
    # TODO: parse the HTML body of the email to plain text
    # for email in email: 
    #     email['body'] = parseHtml(email.get("body"))
    return emails
