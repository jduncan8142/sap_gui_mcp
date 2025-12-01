from logging import getLogger
from typing import Optional
import win32com.client
import os
from datetime import datetime

from sap.logon_pad import sap_session


logger = getLogger("sap_controller")


def export_grid_as_csv(
    grid_id: str,
    session: Optional[win32com.client.CDispatch] = None,
    output_path: Optional[str] = None,
    identifier: Optional[str] = None,
) -> tuple[str, str]:
    """
    Export SAP GUI grid data to a CSV file.

    Args:
        grid_id: SAP GUI element ID of the grid to export.
        output_path: Path where the exported CSV should be saved. If None, generates a default path.
        session: SAP session object. If None, uses the current session.
        identifier: Optional identifier or name for the export operation, such as transaction code or table name. Will be included in the filename if provided.

    Returns:
        tuple[str, str]: String for 'file_path' on success or an empty string on failure and 'error' message if any otherwise an empty string.
    """
    _export_dir = ""
    _export_path = ""
    try:
        # Get current session if not provided
        if not session:
            _current_session = sap_session()
        else:
            _current_session = session

        if not _current_session:
            error_msg = "No current session available. Cannot export data."
            logger.error(error_msg)
            return _export_path, error_msg

        # Generate output path if not provided
        if not output_path:
            # Create export directory in current working directory
            _export_dir = os.path.join(os.getcwd(), "exports")
            os.makedirs(_export_dir, exist_ok=True)
        else:
            # Ensure the directory exists
            _export_dir = os.path.dirname(output_path)
            if _export_dir and not os.path.exists(_export_dir):
                os.makedirs(_export_dir, exist_ok=True)

        # Generate filename with timestamp and identifier if provided
        _identifier_part = f"{identifier}_" if identifier else ""
        _export_path = os.path.join(
            _export_dir,
            f"export_{_identifier_part}{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        )

        # Find the grid control
        _grid = _current_session.FindById(grid_id)
        if not _grid:
            error_msg = f"Grid with ID '{grid_id}' not found."
            logger.error(error_msg)
            return _export_path, error_msg

        # Check that the grid is subtype of GridView
        if _grid.Subtype != "GridView":
            error_msg = f"Element with ID '{grid_id}' is not a GridView. Found type: {_grid.Subtype}"
            logger.error(error_msg)
            return _export_path, error_msg

        # Perform the export
        try:
            _grid.pressToolbarContextButton("&MB_EXPORT")
            _grid.selectContextMenuItem("&XXL")
            _current_session.FindById(
                "wnd[1]/usr/ssubSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME"
            ).text = _export_path
            _current_session.FindById(
                "wnd[1]/usr/ssubSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/cmbGS_EXPORT-FORMAT"
            ).setFocus()
            _current_session.FindById(
                "wnd[1]/usr/ssubSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/cmbGS_EXPORT-FORMAT"
            ).key = "csv-LEAN-STANDARD"
            _current_session.FindById("wnd[1]/tbar[0]/btn[20]").press()
            _current_session.FindById("wnd[1]/tbar[0]/btn[0]").press()
        except Exception as e:
            error_msg = f"Failed to export grid data: {str(e)}"
            logger.error(error_msg)
            return _export_path, error_msg

        # Verify the file was created
        if os.path.exists(_export_path):
            success_msg = f"export saved successfully to: {_export_path}"
            logger.info(success_msg)
            return _export_path, success_msg
        else:
            error_msg = f"Export completed but file not found at: {_export_path}"
            logger.warning(error_msg)
            return _export_path, error_msg

    except Exception as e:
        error_msg = f"Failed to export data: {str(e)}"
        logger.error(error_msg)
        return _export_path, error_msg
