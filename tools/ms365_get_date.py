from datetime import datetime

from ibm_watsonx_orchestrate.agent_builder.tools import tool, ToolPermission


@tool(
    name="ms365_get_date",
    description="Get current date",
    permission=ToolPermission.READ_ONLY,
)
def list_calendar_events() -> str:
    """
    Get current date

    Returns:
        str: Current date, in ISO format.
    """
    return datetime.now().isoformat()
