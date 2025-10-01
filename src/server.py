from fastmcp import FastMCP
import logging
from typing import Optional
import json
import win32com.client
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Initialize FastMCP server
mcp = FastMCP("SAP GUI MCP Server")

logger = logging.getLogger(__name__)


def sap_session() -> Optional[win32com.client.CDispatch]:
    """Get the current SAP session."""
    sap_gui = win32com.client.GetObject("SAPGUI")
    if not sap_gui:
        logger.error("SAP GUI is not running.")
        return None
    sap_application = sap_gui.GetScriptingEngine
    if not sap_application.Connections:
        logger.error("No SAP connections found.")
        return None
    connection = sap_application.Connections[0]
    if not connection.Sessions:
        logger.error("No SAP sessions found.")
        return None
    return connection.Sessions[0]


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


@mcp.tool()
def get_session_info() -> str:
    """Get information about the current SAP session."""
    _current_session = sap_session()
    if not _current_session:
        return "No current session available."
    try:
        info = {
            "SessionId": _current_session.Id,
            "User": _current_session.Info.User,
            "Client": _current_session.Info.Client,
            "Language": _current_session.Info.Language,
            "SystemName": _current_session.Info.SystemName,
            "SystemNumber": _current_session.Info.SystemNumber,
        }
        logger.info("Retrieved session information.")
        return json.dumps(info, indent=2)
    except Exception as e:
        error_msg = f"Failed to retrieve session information: {str(e)}"
        logger.error(error_msg)
        return error_msg


@mcp.tool()
def start_transaction(transaction_code: str) -> str:
    """Start a new SAP transaction by its transaction code."""
    if not transaction_code:
        error_msg = "No transaction code specified."
        logger.error(error_msg)
        return error_msg
    _current_session = sap_session()
    if not _current_session:
        error_msg = "No current session available. Cannot start transaction."
        logger.error(error_msg)
        return error_msg
    try:
        _current_session.StartTransaction(transaction_code)
        logger.info(f"Started transaction: {transaction_code}")
        return f"Started transaction: {transaction_code}"
    except Exception as e:
        error_msg = f"Failed to start transaction '{transaction_code}': {str(e)}"
        logger.error(error_msg)
        return error_msg


@mcp.tool()
def end_transaction() -> str:
    """End the current SAP transaction."""
    _current_session = sap_session()
    if not _current_session:
        error_msg = "No current session available. Cannot end transaction."
        logger.error(error_msg)
        return error_msg
    try:
        _current_session.EndTransaction()
        logger.info("Ended current transaction.")
        return "Ended current transaction."
    except Exception as e:
        error_msg = f"Failed to end transaction: {str(e)}"
        logger.error(error_msg)
        return error_msg


@mcp.tool()
def get_sap_gui_tree() -> str:
    """Get a textual representation of the current SAP GUI tree."""
    _current_session = sap_session()
    if not _current_session:
        error_msg = "No current session available. Cannot retrieve GUI tree."
        logger.error(error_msg)
        return error_msg
    try:
        # _gui_tree = json.dumps(json.loads(_current_session.GetObjectTree(_current_session.Children[0].Id)), indent=2)
        _gui_tree = json.dumps(sap_object_tree_as_json(session=_current_session), indent=2)
        logger.info("Retrieved SAP GUI tree.")
        return _gui_tree
    except Exception as e:
        error_msg = f"Failed to retrieve SAP GUI tree: {str(e)}"
        logger.error(error_msg)
        return error_msg


@mcp.tool()
def find_by_id(element_id: str, raise_error: Optional[bool] = False) -> str:
    """Find a GUI element by its ID."""
    if not element_id:
        error_msg = "No element ID specified."
        logger.error(error_msg)
        return error_msg
    _current_session = sap_session()
    if not _current_session:
        error_msg = "No current session available. Cannot find element."
        logger.error(error_msg)
        return error_msg
    try:
        _current_element = _current_session.FindById(element_id, raise_error)
        if _current_element:
            logger.info(f"Found element by ID: {element_id}")
            return f"Found element {_current_element.Id} by ID: {element_id}"
        else:
            error_msg = f"No element found with ID: {element_id}"
            logger.error(error_msg)
            return error_msg
    except Exception as e:
        error_msg = f"Failed to find element with ID '{element_id}': {str(e)}"
        logger.error(error_msg)
        return error_msg


@mcp.tool()
def send_command(command: str) -> str:
    """Send a command to the current SAP session."""
    if not command:
        error_msg = "No command specified."
        logger.error(error_msg)
        return error_msg
    _current_session = sap_session()
    if not _current_session:
        error_msg = "No current session available. Cannot send command."
        logger.error(error_msg)
        return error_msg
    try:
        _current_session.SendCommand(command)
        logger.info(f"Sent command: {command}")
        return f"Sent command: {command}"
    except Exception as e:
        error_msg = f"Failed to send command '{command}': {str(e)}"
        logger.error(error_msg)
        return error_msg


@mcp.tool()
def send_command_async(command: str) -> str:
    """Send a command asynchronously to the current SAP session."""
    if not command:
        error_msg = "No command specified."
        logger.error(error_msg)
        return error_msg
    _current_session = sap_session()
    if not _current_session:
        error_msg = "No current session available. Cannot send command."
        logger.error(error_msg)
        return error_msg
    try:
        _current_session.SendCommandAsync(command)
        logger.info(f"Sent command asynchronously: {command}")
        return f"Sent command asynchronously: {command}"
    except Exception as e:
        error_msg = f"Failed to send command '{command}' asynchronously: {str(e)}"
        logger.error(error_msg)
        return error_msg


@mcp.tool()
def check_gui_busy() -> str:
    """Check if the SAP GUI is busy."""
    _current_session = sap_session()
    if not _current_session:
        error_msg = "No current session available. Cannot check if GUI is busy."
        logger.error(error_msg)
        return error_msg
    try:
        is_busy = _current_session.Busy
        logger.info(f"SAP GUI busy status: {is_busy}")
        return f"SAP GUI busy status: {is_busy}"
    except Exception as e:
        error_msg = f"Failed to check if SAP GUI is busy: {str(e)}"
        logger.error(error_msg)
        return error_msg


@mcp.tool()
def find_by_name(element_name: str, element_type: str) -> str:
    """Find a GUI element by its name and type."""
    if not element_name:
        error_msg = "No element name specified."
        logger.error(error_msg)
        return error_msg
    _current_session = sap_session()
    if not _current_session:
        error_msg = "No current session available. Cannot find element."
        logger.error(error_msg)
        return error_msg
    try:
        _current_element = _current_session.FindByName(element_name, element_type)
        if _current_element:
            logger.info(f"Found element by name: {element_name}")
            return f"Found element {_current_element.Id} by name: {element_name}"
        else:
            error_msg = f"No element found with name: {element_name}"
            logger.error(error_msg)
            return error_msg
    except Exception as e:
        error_msg = f"Failed to find element with name '{element_name}': {str(e)}"
        logger.error(error_msg)
        return error_msg


@mcp.tool()
def find_all_by_name(element_name: str, element_type: str) -> list[str] | str:
    """Find all GUI elements by their name and type."""
    if not element_name:
        error_msg = "No element name specified."
        logger.error(error_msg)
        return error_msg
    _current_session = sap_session()
    if not _current_session:
        error_msg = "No current session available. Cannot find elements."
        logger.error(error_msg)
        return error_msg
    try:
        _elements = _current_session.FindAllByName(element_name, element_type)
        if _elements:
            element_ids = [elem.Id for elem in _elements]
            logger.info(f"Found {len(_elements)} elements by name: {element_name}")
            return element_ids
        else:
            error_msg = f"No elements found with name: {element_name}"
            logger.error(error_msg)
            return error_msg
    except Exception as e:
        error_msg = f"Failed to find elements with name '{element_name}': {str(e)}"
        logger.error(error_msg)
        return error_msg


@mcp.tool()
def set_text(element_id: str, text: str) -> str:
    """Set text in a GUI element by its ID."""
    if not element_id:
        error_msg = "No element ID specified."
        logger.error(error_msg)
        return error_msg
    _current_session = sap_session()
    if not _current_session:
        error_msg = "No current session available. Cannot set text."
        logger.error(error_msg)
        return error_msg
    try:
        _element = _current_session.FindById(element_id)
        if not _element:
            error_msg = f"No element found with ID: {element_id}"
            logger.error(error_msg)
            return error_msg
        _element.Text = text
        logger.info(f"Set text for element ID {element_id} to '{text}'")
        return f"Set text for element ID {element_id} to '{text}'"
    except Exception as e:
        error_msg = f"Failed to set text for element ID '{element_id}': {str(e)}"
        logger.error(error_msg)
        return error_msg


@mcp.tool()
def get_text(element_id: str) -> str:
    """Get text from a GUI element by its ID."""
    if not element_id:
        error_msg = "No element ID specified."
        logger.error(error_msg)
        return error_msg
    _current_session = sap_session()
    if not _current_session:
        error_msg = "No current session available. Cannot get text."
        logger.error(error_msg)
        return error_msg
    try:
        _element = _current_session.FindById(element_id)
        if not _element:
            error_msg = f"No element found with ID: {element_id}"
            logger.error(error_msg)
            return error_msg
        text = _element.Text
        logger.info(f"Got text from element ID {element_id}: '{text}'")
        return text
    except Exception as e:
        error_msg = f"Failed to get text for element ID '{element_id}': {str(e)}"
        logger.error(error_msg)
        return error_msg


@mcp.tool()
def press_button(element_id: str) -> str:
    """Press a button in the GUI by its element ID."""
    if not element_id:
        error_msg = "No element ID specified."
        logger.error(error_msg)
        return error_msg
    _current_session = sap_session()
    if not _current_session:
        error_msg = "No current session available. Cannot press button."
        logger.error(error_msg)
        return error_msg
    try:
        _element = _current_session.FindById(element_id)
        if not _element:
            error_msg = f"No element found with ID: {element_id}"
            logger.error(error_msg)
            return error_msg
        _element.Press()
        logger.info(f"Pressed button with element ID: {element_id}")
        return f"Pressed button with element ID: {element_id}"
    except Exception as e:
        error_msg = f"Failed to press button with element ID '{element_id}': {str(e)}"
        logger.error(error_msg)
        return error_msg


@mcp.tool()
def set_radio_button(element_id: str, selected: bool) -> str:
    """Set a radio button's selected state by its element ID."""
    if not element_id:
        error_msg = "No element ID specified."
        logger.error(error_msg)
        return error_msg
    _current_session = sap_session()
    if not _current_session:
        error_msg = "No current session available. Cannot set radio button."
        logger.error(error_msg)
        return error_msg
    try:
        _element = _current_session.FindById(element_id)
        if not _element:
            error_msg = f"No element found with ID: {element_id}"
            logger.error(error_msg)
            return error_msg
        _element.Select()
        logger.info(f"Set radio button with element ID {element_id} to selected={selected}")
        return f"Set radio button with element ID {element_id} to selected={selected}"
    except Exception as e:
        error_msg = f"Failed to set radio button with element ID '{element_id}': {str(e)}"
        logger.error(error_msg)
        return error_msg


@mcp.tool()
def get_grid_data(element_id: str) -> dict:
    """
    Extract all data from an SAP GUI grid control.

    Args:
        element_id: The ID of the grid shell element (e.g., '/app/con[0]/ses[0]/wnd[0]/usr/cntlGRID1/shellcont/shell')

    Returns:
        Dictionary containing grid data with columns and rows
    """
    try:
        session = sap_session()
        if not session:
            return {"error": "No current session available."}
        grid = session.findById(element_id)

        # Get grid dimensions
        row_count = grid.RowCount
        visible_row_count = grid.VisibleRowCount
        column_count = grid.ColumnCount

        # Get column information
        column_order = grid.ColumnOrder
        columns = [grid.GetColumnTitles(column_order.Item(col_idx)).Item(0) for col_idx in range(column_count - 1)]

        # Extract row data
        rows = []
        for row_idx in range(row_count):
            row_data = {}
            for col_idx in range(column_count - 1):
                cell_value = grid.GetCellValue(row_idx, column_order.Item(col_idx))
                row_data[column_order.Item(col_idx)] = {"Column": columns[col_idx], "Value": cell_value}
            rows.append({row_idx: row_data})

        return {
            "row_count": row_count,
            "visible_row_count": visible_row_count,
            "column_count": column_count,
            "columns": columns,
            "rows": rows,
        }

    except Exception as e:
        return {"error": str(e), "message": f"Failed to extract grid data from element: {element_id}"}


@mcp.tool()
def get_grid_cell_value(element_id: str, row: int, column: int) -> dict:
    """
    Get a specific cell value from an SAP GUI grid control.

    Args:
        element_id: The ID of the grid shell element
        row: Row index (0-based)
        column: Column index (0-based)

    Returns:
        Dictionary containing the cell value
    """
    try:
        session = sap_session()
        if not session:
            return {"error": "No current session available."}
        grid = session.findById(element_id)

        cell_value = grid.GetCellValue(row, column)
        column_title = grid.GetColumnTitle(column)

        return {"row": row, "column": column, "column_title": column_title, "value": cell_value}

    except Exception as e:
        return {"error": str(e), "message": f"Failed to get cell value at row {row}, column {column}"}


@mcp.tool()
def select_grid_row(element_id: str, row: int) -> dict:
    """
    Select a specific row in an SAP GUI grid control.

    Args:
        element_id: The ID of the grid shell element
        row: Row index (0-based)

    Returns:
        Dictionary confirming the selection
    """
    try:
        session = sap_session()
        if not session:
            return {"error": "No current session available."}
        grid = session.findById(element_id)

        # Set current cell position to select the row
        grid.CurrentCellRow = row
        grid.SelectedRows = str(row)

        return {"success": True, "selected_row": row, "message": f"Selected row {row}"}

    except Exception as e:
        return {"error": str(e), "message": f"Failed to select row {row}"}


@mcp.tool()
def get_selected_grid_rows(element_id: str) -> dict:
    """
    Get data from currently selected rows in an SAP GUI grid control.

    Args:
        element_id: The ID of the grid shell element

    Returns:
        Dictionary containing selected rows data
    """
    try:
        session = sap_session()
        if not session:
            return {"error": "No current session available."}
        grid = session.findById(element_id)

        # Get selected rows
        selected_rows_str = grid.SelectedRows

        if not selected_rows_str:
            return {"selected_row_count": 0, "rows": []}

        # Parse selected rows (format can be "1,2,3" or "1-5")
        selected_indices = []
        for part in selected_rows_str.split(","):
            if "-" in part:
                start, end = map(int, part.split("-"))
                selected_indices.extend(range(start, end + 1))
            else:
                selected_indices.append(int(part))

        # Get column information
        column_count = grid.ColumnCount
        column_names = [grid.GetColumnTitle(i) for i in range(column_count)]

        # Extract data from selected rows
        rows = []
        for row_idx in selected_indices:
            row_data = {}
            for col_idx in range(column_count):
                cell_value = grid.GetCellValue(row_idx, col_idx)
                row_data[column_names[col_idx]] = cell_value
            rows.append(row_data)

        return {"selected_row_count": len(selected_indices), "selected_indices": selected_indices, "rows": rows}

    except Exception as e:
        return {"error": str(e), "message": "Failed to get selected rows"}


@mcp.tool()
def double_click_grid_cell(element_id: str, row: int, column: int) -> dict:
    """
    Double-click on a specific cell in an SAP GUI grid control.
    This typically opens the detail view or drills down into the record.

    Args:
        element_id: The ID of the grid shell element
        row: Row index (0-based)
        column: Column index (0-based)

    Returns:
        Dictionary confirming the action
    """
    try:
        session = sap_session()
        if not session:
            return {"error": "No current session available."}
        grid = session.findById(element_id)

        # Set current cell and double-click
        grid.CurrentCellRow = row
        grid.CurrentCellColumn = column
        grid.DoubleClick(row, column)

        return {
            "success": True,
            "row": row,
            "column": column,
            "message": f"Double-clicked cell at row {row}, column {column}",
        }

    except Exception as e:
        return {"error": str(e), "message": f"Failed to double-click cell at row {row}, column {column}"}


@mcp.tool()
def set_checkbox(element_id: str, state: bool) -> str:
    """Set the state of a checkbox element by its ID."""
    if not element_id:
        error_msg = "No element ID specified."
        logger.error(error_msg)
        return error_msg
    _current_session = sap_session()
    if not _current_session:
        error_msg = "No current session available. Cannot set checkbox."
        logger.error(error_msg)
        return error_msg
    try:
        _element = _current_session.FindById(element_id)
        if not _element:
            error_msg = f"No element found with ID: {element_id}"
            logger.error(error_msg)
            return error_msg
        _element.Selected = state
        logger.info(f"Set checkbox with element ID {element_id} to state={state}")
        return f"Set checkbox with element ID {element_id} to state={state}"
    except Exception as e:
        error_msg = f"Failed to set checkbox with element ID '{element_id}': {str(e)}"
        logger.error(error_msg)
        return error_msg


@mcp.tool()
def get_vertical_scrollbar_position(element_id: str) -> dict:
    """
    Get the current position of the vertical scrollbar in an SAP GUI grid control.

    Args:
        element_id: The ID of the grid shell element

    Returns:
        Dictionary containing the vertical scrollbar position
    """
    if not element_id:
        return {"error": "No element ID specified."}

    session = sap_session()
    if not session:
        return {"error": "No current session available."}

    try:
        grid = session.findById(element_id)
        if not grid:
            return {"error": f"No grid found with ID: {element_id}"}

        scrollbar_position = grid.VerticalScrollbar.Position
        return {"success": True, "scrollbar_position": scrollbar_position}

    except Exception as e:
        return {"error": str(e), "message": "Failed to get vertical scrollbar position"}


@mcp.tool()
def set_vertical_scrollbar_position(element_id: str, position: int) -> dict:
    """
    Set the position of the vertical scrollbar in an SAP GUI grid control.

    Args:
        element_id: The ID of the grid shell element
        position: The desired scrollbar position

    Returns:
        Dictionary indicating success or failure
    """
    if not element_id:
        return {"error": "No element ID specified."}

    session = sap_session()
    if not session:
        return {"error": "No current session available."}

    try:
        grid = session.findById(element_id)
        if not grid:
            return {"error": f"No grid found with ID: {element_id}"}

        grid.VerticalScrollbar.Position = position
        return {"success": True}

    except Exception as e:
        return {"error": str(e), "message": "Failed to set vertical scrollbar position"}


def get_horizontal_scrollbar_position(element_id: str) -> dict:
    """
    Get the current position of the horizontal scrollbar in an SAP GUI grid control.

    Args:
        element_id: The ID of the grid shell element

    Returns:
        Dictionary containing the horizontal scrollbar position
    """
    if not element_id:
        return {"error": "No element ID specified."}

    session = sap_session()
    if not session:
        return {"error": "No current session available."}

    try:
        grid = session.findById(element_id)
        if not grid:
            return {"error": f"No grid found with ID: {element_id}"}

        scrollbar_position = grid.HorizontalScrollbar.Position
        return {"success": True, "scrollbar_position": scrollbar_position}

    except Exception as e:
        return {"error": str(e), "message": "Failed to get horizontal scrollbar position"}


@mcp.tool()
def set_horizontal_scrollbar_position(element_id: str, position: int) -> dict:
    """
    Set the position of the horizontal scrollbar in an SAP GUI grid control.

    Args:
        element_id: The ID of the grid shell element
        position: The desired scrollbar position

    Returns:
        Dictionary indicating success or failure
    """
    if not element_id:
        return {"error": "No element ID specified."}

    session = sap_session()
    if not session:
        return {"error": "No current session available."}

    try:
        grid = session.findById(element_id)
        if not grid:
            return {"error": f"No grid found with ID: {element_id}"}

        grid.HorizontalScrollbar.Position = position
        return {"success": True}

    except Exception as e:
        return {"error": str(e), "message": "Failed to set horizontal scrollbar position"}


@mcp.tool()
def maximize_window() -> str:
    """Maximize the current SAP GUI window."""
    _current_session = sap_session()
    if not _current_session:
        error_msg = "No current session available. Cannot maximize window."
        logger.error(error_msg)
        return error_msg
    try:
        _current_window = _current_session.ActiveWindow
        if not _current_window:
            error_msg = "No active window found in the current session."
            logger.error(error_msg)
            return error_msg
        _current_window.Maximize()
        logger.info("Maximized the current SAP GUI window.")
        return "Maximized the current SAP GUI window."
    except Exception as e:
        error_msg = f"Failed to maximize window: {str(e)}"
        logger.error(error_msg)
        return error_msg


@mcp.tool()
def set_focus(element_id: str) -> str:
    """Set focus to a GUI element by its ID."""
    if not element_id:
        error_msg = "No element ID specified."
        logger.error(error_msg)
        return error_msg
    _current_session = sap_session()
    if not _current_session:
        error_msg = "No current session available. Cannot set focus."
        logger.error(error_msg)
        return error_msg
    try:
        _element = _current_session.FindById(element_id)
        if not _element:
            error_msg = f"No element found with ID: {element_id}"
            logger.error(error_msg)
            return error_msg
        _element.SetFocus()
        logger.info(f"Set focus to element with ID: {element_id}")
        return f"Set focus to element with ID: {element_id}"
    except Exception as e:
        error_msg = f"Failed to set focus for element ID '{element_id}': {str(e)}"
        logger.error(error_msg)
        return error_msg


@mcp.tool()
def set_combobox(element_id: str, key: str) -> str:
    """Set the value of a combo box element by its ID."""
    if not element_id:
        error_msg = "No element ID specified."
        logger.error(error_msg)
        return error_msg
    _current_session = sap_session()
    if not _current_session:
        error_msg = "No current session available. Cannot set combo box."
        logger.error(error_msg)
        return error_msg
    try:
        _element = _current_session.FindById(element_id)
        if not _element:
            error_msg = f"No element found with ID: {element_id}"
            logger.error(error_msg)
            return error_msg
        _element.Key = key
        logger.info(f"Set combo box with element ID {element_id} to key='{key}'")
        return f"Set combo box with element ID {element_id} to key='{key}'"
    except Exception as e:
        error_msg = f"Failed to set combo box with element ID '{element_id}': {str(e)}"
        logger.error(error_msg)
        return error_msg


if __name__ == "__main__":
    mcp.run(transport="stdio")
