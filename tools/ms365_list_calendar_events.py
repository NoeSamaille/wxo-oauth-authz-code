import requests
from typing import Optional
from datetime import datetime, timedelta
from pydantic import Field, BaseModel

from ibm_watsonx_orchestrate.agent_builder.tools import tool, ToolPermission
from ibm_watsonx_orchestrate.run import connections
from ibm_watsonx_orchestrate.agent_builder.connections import ConnectionType


CONNECTION_MS365 = "ms365"


class Event(BaseModel):
    """
    Represents a calendar event.
    """

    subject: str = Field(..., description="The subject of the event")
    start: str = Field(..., description="The time at which the event starts")
    end: str = Field(None, description="The time at which the event ends")
    description: Optional[str] = Field(
        None, description="Detailed information about the event"
    )


@tool(
    name="ms365_list_calendar_events",
    description="Get Microsoft 365 calendar events",
    permission=ToolPermission.READ_ONLY,
    expected_credentials=[
        {"app_id": CONNECTION_MS365, "type": ConnectionType.OAUTH2_AUTH_CODE}
    ],
)
def list_calendar_events(
    start: Optional[str] = None, offset_day: int = 5
) -> list:
    """
    Get Microsoft 365 calendar events.
    
    Args:
        start (string): Optional, the date to start the search, format 'YYYY-MM-DD, defaults to the current date if None.
        offset_day (int): The offset in days to end the search, defaults to 5, minimum is 1.

    Returns:
        list: A list of calendar events.
    """
    conn = connections.oauth2_auth_code(CONNECTION_MS365)
    given_date = datetime.fromisoformat(start)
    start_datetime = (
        given_date.replace(hour=0, minute=0, second=0, microsecond=0).isoformat() + "Z"
    )
    end_datetime = (given_date + timedelta(days=offset_day)).replace(
        hour=0, minute=0, second=0, microsecond=0
    ).isoformat() + "Z"

    # Graph API endpoint for calendar events in the specified range
    url = f"https://graph.microsoft.com/v1.0/me/calendarView?startDateTime={start_datetime}&endDateTime={end_datetime}"

    # Set headers with the OAuth2 access token
    headers = {
        "Authorization": f"Bearer {conn.access_token}",
        "Prefer": 'outlook.timezone="UTC"',
    }

    # Make the API request
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    events = response.json().get("value", [])

    event_list = []
    print(f"Events on {given_date.date()}:")
    for event in events:
        event_model = Event(
            subject=event.get("subject", ""),
            start=event.get("start", {}).get("dateTime", ""),
            end=event.get("end", {}).get("dateTime", ""),
        )
        event_list.append(event_model)
        print(
            "-",
            event.get("subject"),
            "|",
            event.get("start", {}).get("dateTime"),
            "-",
            event.get("end", {}).get("dateTime"),
        )
    return event_list
