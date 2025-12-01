from logging import getLogger
from typing import Optional
import win32com.client
import json
import os
from datetime import datetime

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


def capture_screenshot(
    output_path: Optional[str] = None,
    window_id: Optional[str] = None,
    session: Optional[win32com.client.CDispatch] = None,
) -> tuple[bool, str]:
    """
    Capture a screenshot of the SAP GUI window.

    Args:
        output_path: Path where the screenshot should be saved. If None, generates a default path.
        window_id: ID of the window to capture (e.g., "wnd[0]"). If None, captures the active window.
        session: SAP session object. If None, uses the current session.

    Returns:
        Tuple of (success: bool, message: str) where message contains the file path on success or error message on failure
    """
    try:
        # Get current session if not provided
        if not session:
            _current_session = sap_session()
        else:
            _current_session = session

        if not _current_session:
            error_msg = "No current session available. Cannot capture screenshot."
            logger.error(error_msg)
            return False, error_msg

        # Determine which window to capture
        if window_id:
            try:
                window = _current_session.FindById(window_id)
                if not window:
                    error_msg = f"Window with ID '{window_id}' not found."
                    logger.error(error_msg)
                    return False, error_msg
            except Exception as e:
                error_msg = f"Failed to find window '{window_id}': {str(e)}"
                logger.error(error_msg)
                return False, error_msg
        else:
            # Use active window
            window = _current_session.ActiveWindow
            if not window:
                error_msg = "No active window found in the current session."
                logger.error(error_msg)
                return False, error_msg

        # Generate output path if not provided
        if not output_path:
            # Create screenshots directory in current working directory
            screenshots_dir = os.path.join(os.getcwd(), "screenshots")
            os.makedirs(screenshots_dir, exist_ok=True)

            # Generate filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(screenshots_dir, f"sap_screenshot_{timestamp}.png")
        else:
            # Ensure the directory exists
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)

        # Capture the screenshot using SAP GUI's HardCopy method
        # HardCopy saves the window as an image file
        # The method signature is: HardCopy(filename, format)
        # Format: 0 = BMP, 1 = DIB, 2 = JPEG, 3 = PNG
        window.HardCopy(output_path, 3)  # 3 = PNG format

        # Verify the file was created
        if os.path.exists(output_path):
            success_msg = f"Screenshot saved successfully to: {output_path}"
            logger.info(success_msg)
            return True, success_msg
        else:
            error_msg = f"Screenshot capture completed but file not found at: {output_path}"
            logger.warning(error_msg)
            return False, error_msg

    except Exception as e:
        error_msg = f"Failed to capture screenshot: {str(e)}"
        logger.error(error_msg)
        return False, error_msg
