from logging import getLogger
from typing import Optional
import win32com.client
import json

from sap.logon_pad import sap_session

logger = getLogger("sap_controller")
MAX_RETRIES = 10
WAIT_TIME = 3.0
POLLING_INTERVAL = 0.5


def sap_object_tree_as_json(session: Optional[win32com.client.CDispatch]) -> dict:
    """Convert SAP object tree string to JSON."""
    try:
        _json_objects = []
        if not session:
            _current_session = sap_session()
        else:
            _current_session = session
        if not _current_session:
            logger.error("No current session available.")
            return {"Windows": []}
        _windows = _current_session.Children
        for window in _windows:
            _win = json.loads(_current_session.GetObjectTree(window.Id))
            _json_objects.append(_win)
        return {"Windows": _json_objects}
    except json.JSONDecodeError as e:
        logger.error(f"Failed to decode SAP object tree: {str(e)}")
        return {"Windows": []}
