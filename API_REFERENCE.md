# SAP GUI MCP - API Reference

Complete reference for all available MCP tools in the SAP GUI MCP Server.

## Table of Contents

- [Session Management](#session-management)
- [Navigation](#navigation)
- [Element Discovery](#element-discovery)
- [Input/Output Operations](#inputoutput-operations)
- [Command Execution](#command-execution)
- [Grid Controls](#grid-controls)
- [Window Management](#window-management)
- [Return Value Formats](#return-value-formats)

---

## Session Management

### `get_session_info()`

**Description**: Retrieves information about the current SAP session.

**Parameters**: None

**Returns**: JSON string containing session information

**Return Structure**:
```json
{
  "SessionId": "ses[0]",
  "User": "USERNAME",
  "Client": "100",
  "Language": "EN",
  "SystemName": "PRD",
  "SystemNumber": "00"
}
```

**Example**:
```python
session_info = get_session_info()
# Returns: JSON string with session details
```

**Error Returns**:
- `"No current session available."` - When SAP GUI is not running or no session exists
- `"Failed to retrieve session information: <error>"` - When COM error occurs

**Source**: `src/server.py:212`

---

### `login_to_sap()` ⭐ NEW

**Description**: Creates a new SAP session by logging in with credentials. Enables automatic session creation without manual SAP GUI login.

**Parameters**:
- `system` (str, optional): SAP system ID (e.g., "PRD", "DEV", "QAS"). If None, uses `SAP_SYSTEM` environment variable
- `client` (str, optional): SAP client number (e.g., "100", "200"). If None, uses `SAP_CLIENT` environment variable
- `user` (str, optional): SAP username. If None, uses `SAP_USER` environment variable
- `password` (str, optional): SAP password. If None, uses `SAP_PASSWORD` environment variable
- `language` (str, optional): SAP language code (default: "EN"). If None, uses `SAP_LANGUAGE` environment variable

**Returns**: JSON string containing session information or error message

**Return Structure** (Success):
```json
{
  "Success": true,
  "SessionId": "ses[0]",
  "User": "USERNAME",
  "Client": "100",
  "Language": "EN",
  "SystemName": "PRD",
  "SystemNumber": "00"
}
```

**Example 1** (Using environment variables):
```python
# Requires .env file with SAP credentials configured
result = login_to_sap()
# Returns: Session info JSON
```

**Example 2** (Using explicit credentials):
```python
result = login_to_sap(
    system="PRD",
    client="100",
    user="MYUSER",
    password="MyPass123",
    language="EN"
)
# Returns: Session info JSON
```

**Example 3** (Mixing environment and explicit):
```python
# Use env for most, override password only
result = login_to_sap(password="DifferentPassword")
# Uses SAP_SYSTEM, SAP_CLIENT, SAP_USER from .env, but provided password
```

**Error Returns**:
- `"Login failed: Missing required parameters: <list>"` - When required credentials are not provided
- `"Failed to create SAP session for system <system>..."` - When login fails (wrong credentials, system unavailable, etc.)
- `"SAP GUI is not running. Please start SAP GUI first."` - When SAP GUI application is not running
- `"Failed to open connection to system: <system>"` - When connection to SAP system fails

**Prerequisites**:
1. SAP GUI for Windows must be running
2. SAP system must be configured in SAP Logon
3. Valid credentials (either in `.env` or passed as parameters)
4. SAP GUI Scripting must be enabled

**Use Cases**:
- Automated session creation for CI/CD pipelines
- Recover from disconnected sessions
- Create sessions without manual login
- Test different user accounts programmatically

**Source**: `src/server.py:234`

---

### `check_gui_busy()`

**Description**: Checks if the SAP GUI is currently processing a request.

**Parameters**: None

**Returns**: String indicating busy status

**Example**:
```python
status = check_gui_busy()
# Returns: "SAP GUI busy status: True" or "SAP GUI busy status: False"
```

**Use Case**: Wait for SAP to complete processing before next action

**Source**: `src/server.py:209`

---

## Navigation

### `start_transaction(transaction_code: str)`

**Description**: Starts a new SAP transaction by its transaction code.

**Parameters**:
- `transaction_code` (str, required): The SAP transaction code (e.g., "SE38", "VA01")

**Returns**: Confirmation string

**Example**:
```python
result = start_transaction("SE38")
# Returns: "Started transaction: SE38"
```

**Common Transaction Codes**:
- `SE38` - ABAP Editor
- `VA01` - Create Sales Order
- `MM01` - Create Material Master
- `FB01` - Post Document
- `SE16` - Data Browser

**Error Returns**:
- `"No transaction code specified."` - When empty/None provided
- `"No current session available. Cannot start transaction."` - No SAP session
- `"Failed to start transaction '<code>': <error>"` - Transaction error

**Source**: `src/server.py:79`

---

### `end_transaction()`

**Description**: Ends the current SAP transaction.

**Parameters**: None

**Returns**: Confirmation string

**Example**:
```python
result = end_transaction()
# Returns: "Ended current transaction."
```

**Note**: Equivalent to pressing `/n` in SAP GUI

**Source**: `src/server.py:101`

---

## Element Discovery

### `get_sap_gui_tree()`

**Description**: Retrieves a JSON representation of the current SAP GUI object tree.

**Parameters**: None

**Returns**: JSON string containing the complete GUI tree structure

**Return Structure**:
```json
{
  "Windows": [
    {
      "id": "wnd[0]",
      "type": "GuiMainWindow",
      "children": [
        {
          "id": "wnd[0]/usr",
          "type": "GuiUserArea",
          "children": [...]
        }
      ]
    }
  ]
}
```

**Example**:
```python
tree = get_sap_gui_tree()
# Returns: Complete JSON tree of GUI elements
```

**Use Case**: Essential for discovering element IDs before interaction

**Source**: `src/server.py:119`

---

### `find_by_id(element_id: str, raise_error: bool = False)`

**Description**: Finds a GUI element by its ID.

**Parameters**:
- `element_id` (str, required): The element ID (e.g., "wnd[0]/usr/txtRF02D-KUNNR")
- `raise_error` (bool, optional): Whether to raise error if not found (default: False)

**Returns**: Confirmation string with element ID

**Example**:
```python
result = find_by_id("wnd[0]/usr/txtRF02D-KUNNR")
# Returns: "Found element wnd[0]/usr/txtRF02D-KUNNR by ID: wnd[0]/usr/txtRF02D-KUNNR"
```

**Element ID Format**:
- `wnd[X]` - Window index
- `/usr` - User area
- `/txt` - Text field
- `/btn` - Button
- `/chk` - Checkbox
- `/cmb` - Combobox

**Source**: `src/server.py:138`

---

### `find_by_name(element_name: str, element_type: str)`

**Description**: Finds a GUI element by its name and type.

**Parameters**:
- `element_name` (str, required): The element name
- `element_type` (str, required): The element type (e.g., "GuiTextField", "GuiButton")

**Returns**: Confirmation string with element ID

**Example**:
```python
result = find_by_name("RF02D-KUNNR", "GuiTextField")
# Returns: "Found element <id> by name: RF02D-KUNNR"
```

**Common Element Types**:
- `GuiTextField` - Text input field
- `GuiButton` - Button
- `GuiCheckBox` - Checkbox
- `GuiComboBox` - Dropdown/Combobox
- `GuiRadioButton` - Radio button
- `GuiGridView` - Grid/ALV control

**Source**: `src/server.py:227`

---

### `find_all_by_name(element_name: str, element_type: str)`

**Description**: Finds all GUI elements matching the given name and type.

**Parameters**:
- `element_name` (str, required): The element name
- `element_type` (str, required): The element type

**Returns**: List of element IDs (or error string)

**Example**:
```python
elements = find_all_by_name("BUTTON", "GuiButton")
# Returns: ["wnd[0]/usr/btn1", "wnd[0]/usr/btn2", ...]
```

**Use Case**: Finding all buttons, checkboxes, or repeated fields

**Source**: `src/server.py:254`

---

## Input/Output Operations

### `set_text(element_id: str, text: str)`

**Description**: Sets text in a GUI element.

**Parameters**:
- `element_id` (str, required): The element ID
- `text` (str, required): The text to set

**Returns**: Confirmation string

**Example**:
```python
result = set_text("wnd[0]/usr/txtRF02D-KUNNR", "12345")
# Returns: "Set text for element ID wnd[0]/usr/txtRF02D-KUNNR to '12345'"
```

**Source**: `src/server.py:282`

---

### `get_text(element_id: str)`

**Description**: Retrieves text from a GUI element.

**Parameters**:
- `element_id` (str, required): The element ID

**Returns**: The text content of the element

**Example**:
```python
text = get_text("wnd[0]/usr/txtRF02D-KUNNR")
# Returns: "12345"
```

**Source**: `src/server.py:309`

---

### `press_button(element_id: str)`

**Description**: Presses a button in the GUI.

**Parameters**:
- `element_id` (str, required): The button element ID

**Returns**: Confirmation string

**Example**:
```python
result = press_button("wnd[0]/usr/btnEXECUTE")
# Returns: "Pressed button with element ID: wnd[0]/usr/btnEXECUTE"
```

**Common Buttons**:
- Execute button (F8)
- Save button
- Back button
- Cancel button

**Source**: `src/server.py:336`

---

### `set_checkbox(element_id: str, state: bool)`

**Description**: Sets the state of a checkbox element.

**Parameters**:
- `element_id` (str, required): The checkbox element ID
- `state` (bool, required): True to check, False to uncheck

**Returns**: Confirmation string

**Example**:
```python
result = set_checkbox("wnd[0]/usr/chkPARAM-CHECK", True)
# Returns: "Set checkbox with element ID wnd[0]/usr/chkPARAM-CHECK to state=True"
```

**Source**: `src/server.py:580`

---

### `set_radio_button(element_id: str, selected: bool)`

**Description**: Selects a radio button.

**Parameters**:
- `element_id` (str, required): The radio button element ID
- `selected` (bool, required): Selection state (typically True)

**Returns**: Confirmation string

**Example**:
```python
result = set_radio_button("wnd[0]/usr/radP_ONLINE", True)
# Returns: "Set radio button with element ID wnd[0]/usr/radP_ONLINE to selected=True"
```

**Note**: The implementation uses `.Select()` method regardless of the `selected` parameter value

**Source**: `src/server.py:363`

---

### `set_combobox(element_id: str, key: str)`

**Description**: Sets the value of a combo box (dropdown).

**Parameters**:
- `element_id` (str, required): The combobox element ID
- `key` (str, required): The key value to select

**Returns**: Confirmation string

**Example**:
```python
result = set_combobox("wnd[0]/usr/cmbVARI-MONTH", "01")
# Returns: "Set combo box with element ID wnd[0]/usr/cmbVARI-MONTH to key='01'"
```

**Note**: Use the key, not the display text

**Source**: `src/server.py:778`

---

### `set_focus(element_id: str)`

**Description**: Sets focus to a GUI element.

**Parameters**:
- `element_id` (str, required): The element ID

**Returns**: Confirmation string

**Example**:
```python
result = set_focus("wnd[0]/usr/txtRF02D-KUNNR")
# Returns: "Set focus to element with ID: wnd[0]/usr/txtRF02D-KUNNR"
```

**Use Case**: Prepare element for input or trigger element-specific actions

**Source**: `src/server.py:751`

---

## Command Execution

### `send_command(command: str)`

**Description**: Sends a command synchronously to the current SAP session.

**Parameters**:
- `command` (str, required): The SAP command to execute

**Returns**: Confirmation string

**Example**:
```python
result = send_command("/nSE38")
# Returns: "Sent command: /nSE38"
```

**Common Commands**:
- `/n<tcode>` - End transaction and start new one
- `/o<tcode>` - Open transaction in new window
- `/<ok>` - Same as pressing Enter
- `/nend` - Log off

**Source**: `src/server.py:165`

---

### `send_command_async(command: str)`

**Description**: Sends a command asynchronously to the current SAP session.

**Parameters**:
- `command` (str, required): The SAP command to execute

**Returns**: Confirmation string

**Example**:
```python
result = send_command_async("/nSE38")
# Returns: "Sent command asynchronously: /nSE38"
```

**Use Case**: For commands that may take time to complete

**Source**: `src/server.py:187`

---

## Grid Controls

### `get_grid_data(element_id: str)`

**Description**: Extracts all data from an SAP GUI grid control (ALV table).

**Parameters**:
- `element_id` (str, required): The grid shell element ID

**Returns**: Dictionary containing grid data

**Return Structure**:
```json
{
  "row_count": 100,
  "visible_row_count": 20,
  "column_count": 10,
  "columns": ["Column1", "Column2", ...],
  "rows": [
    {
      "0": {
        "COL1": {"Column": "Column1", "Value": "Value1"},
        "COL2": {"Column": "Column2", "Value": "Value2"}
      }
    },
    ...
  ]
}
```

**Example**:
```python
data = get_grid_data("wnd[0]/usr/cntlGRID1/shellcont/shell")
# Returns: Complete grid data with all rows and columns
```

**Use Case**: Extract all visible data from ALV grids for processing

**Source**: `src/server.py:390`

---

### `get_grid_cell_value(element_id: str, row: int, column: int)`

**Description**: Gets a specific cell value from a grid control.

**Parameters**:
- `element_id` (str, required): The grid shell element ID
- `row` (int, required): Row index (0-based)
- `column` (int, required): Column index (0-based)

**Returns**: Dictionary with cell information

**Return Structure**:
```json
{
  "row": 0,
  "column": 1,
  "column_title": "Material Number",
  "value": "100001"
}
```

**Example**:
```python
cell = get_grid_cell_value("wnd[0]/usr/cntlGRID1/shellcont/shell", 0, 1)
# Returns: Cell value at first row, second column
```

**Source**: `src/server.py:437`

---

### `select_grid_row(element_id: str, row: int)`

**Description**: Selects a specific row in a grid control.

**Parameters**:
- `element_id` (str, required): The grid shell element ID
- `row` (int, required): Row index (0-based)

**Returns**: Dictionary confirming selection

**Return Structure**:
```json
{
  "success": true,
  "selected_row": 5,
  "message": "Selected row 5"
}
```

**Example**:
```python
result = select_grid_row("wnd[0]/usr/cntlGRID1/shellcont/shell", 5)
# Returns: Confirmation of row 5 selection
```

**Source**: `src/server.py:465`

---

### `get_selected_grid_rows(element_id: str)`

**Description**: Retrieves data from currently selected rows in a grid.

**Parameters**:
- `element_id` (str, required): The grid shell element ID

**Returns**: Dictionary with selected rows data

**Return Structure**:
```json
{
  "selected_row_count": 2,
  "selected_indices": [5, 7],
  "rows": [
    {
      "Material Number": "100001",
      "Description": "Product A"
    },
    {
      "Material Number": "100002",
      "Description": "Product B"
    }
  ]
}
```

**Example**:
```python
selected = get_selected_grid_rows("wnd[0]/usr/cntlGRID1/shellcont/shell")
# Returns: Data from all selected rows
```

**Note**: Supports both single selections and ranges (e.g., "1,2,3" or "1-5")

**Source**: `src/server.py:493`

---

### `double_click_grid_cell(element_id: str, row: int, column: int)`

**Description**: Double-clicks on a specific cell in a grid control.

**Parameters**:
- `element_id` (str, required): The grid shell element ID
- `row` (int, required): Row index (0-based)
- `column` (int, required): Column index (0-based)

**Returns**: Dictionary confirming action

**Return Structure**:
```json
{
  "success": true,
  "row": 0,
  "column": 1,
  "message": "Double-clicked cell at row 0, column 1"
}
```

**Example**:
```python
result = double_click_grid_cell("wnd[0]/usr/cntlGRID1/shellcont/shell", 0, 1)
# Returns: Confirmation and opens detail view
```

**Use Case**: Navigate to detail screens from grid entries

**Source**: `src/server.py:544`

---

### `get_vertical_scrollbar_position(element_id: str)`

**Description**: Gets the current position of the vertical scrollbar in a grid.

**Parameters**:
- `element_id` (str, required): The grid shell element ID

**Returns**: Dictionary with scrollbar position

**Return Structure**:
```json
{
  "success": true,
  "scrollbar_position": 42
}
```

**Example**:
```python
position = get_vertical_scrollbar_position("wnd[0]/usr/cntlGRID1/shellcont/shell")
# Returns: Current vertical scroll position
```

**Source**: `src/server.py:607`

---

### `set_vertical_scrollbar_position(element_id: str, position: int)`

**Description**: Sets the vertical scrollbar position in a grid.

**Parameters**:
- `element_id` (str, required): The grid shell element ID
- `position` (int, required): The desired scrollbar position

**Returns**: Dictionary indicating success

**Return Structure**:
```json
{
  "success": true
}
```

**Example**:
```python
result = set_vertical_scrollbar_position("wnd[0]/usr/cntlGRID1/shellcont/shell", 0)
# Returns: Success confirmation and scrolls to top
```

**Use Case**: Navigate through large grids programmatically

**Source**: `src/server.py:637`

---

### `set_horizontal_scrollbar_position(element_id: str, position: int)`

**Description**: Sets the horizontal scrollbar position in a grid.

**Parameters**:
- `element_id` (str, required): The grid shell element ID
- `position` (int, required): The desired scrollbar position

**Returns**: Dictionary indicating success

**Return Structure**:
```json
{
  "success": true
}
```

**Example**:
```python
result = set_horizontal_scrollbar_position("wnd[0]/usr/cntlGRID1/shellcont/shell", 0)
# Returns: Success confirmation and scrolls to leftmost
```

**Use Case**: View hidden columns in wide grids

**Source**: `src/server.py:697`

---

### `get_horizontal_scrollbar_position(element_id: str)`

**Description**: Gets the current position of the horizontal scrollbar in a grid.

**Parameters**:
- `element_id` (str, required): The grid shell element ID

**Returns**: Dictionary with scrollbar position

**Return Structure**:
```json
{
  "success": true,
  "scrollbar_position": 10
}
```

**Example**:
```python
position = get_horizontal_scrollbar_position("wnd[0]/usr/cntlGRID1/shellcont/shell")
# Returns: Current horizontal scroll position
```

**⚠️ Note**: This function is defined but **missing the `@mcp.tool()` decorator**, so it's not currently exposed as an MCP tool.

**Source**: `src/server.py:667`

---

## Window Management

### `maximize_window()`

**Description**: Maximizes the current SAP GUI window.

**Parameters**: None

**Returns**: Confirmation string

**Example**:
```python
result = maximize_window()
# Returns: "Maximized the current SAP GUI window."
```

**Use Case**: Ensure full screen visibility for GUI operations

**Source**: `src/server.py:728`

---

## Return Value Formats

### Success Responses

Most tools return strings or dictionaries confirming the action:

**String Format**:
```
"<Action> <details>"
```
Example: `"Started transaction: SE38"`

**Dictionary Format**:
```json
{
  "success": true,
  "message": "Action completed",
  ...additional data...
}
```

### Error Responses

**String Format**:
```
"<Error description>: <details>"
```
Example: `"Failed to start transaction 'SE38': Transaction not found"`

**Dictionary Format**:
```json
{
  "error": "Error message",
  "message": "Detailed description"
}
```

### Common Error Messages

| Error | Meaning | Solution |
|-------|---------|----------|
| `"No current session available."` | SAP GUI not running or no session | Start SAP GUI and log in |
| `"No element ID specified."` | Missing required parameter | Provide element ID |
| `"No element found with ID: <id>"` | Element doesn't exist | Verify element ID using `get_sap_gui_tree()` |
| `"Failed to <action>: <details>"` | COM/SAP error | Check SAP state and element type |

---

## Best Practices

### 1. Always Check Session First
```python
session_info = get_session_info()
# Verify connection before operations
```

### 2. Get GUI Tree Before Element Operations
```python
tree = get_sap_gui_tree()
# Identify correct element IDs
```

### 3. Wait for GUI to Be Ready
```python
status = check_gui_busy()
# Ensure SAP finished processing
```

### 4. Use Appropriate Element Types
- Text fields: `set_text()`
- Buttons: `press_button()`
- Checkboxes: `set_checkbox()`
- Radio buttons: `set_radio_button()`
- Dropdowns: `set_combobox()`

### 5. Handle Grid Data Efficiently
```python
# For full extraction
data = get_grid_data(grid_id)

# For specific cells
cell = get_grid_cell_value(grid_id, row, col)

# For selected rows only
selected = get_selected_grid_rows(grid_id)
```

---

## Quick Reference Table

| Category | Tool | Primary Use |
|----------|------|-------------|
| **Session** | `get_session_info()` | Verify connection |
| | `login_to_sap()` ⭐ | Create session/login |
| | `check_gui_busy()` | Wait for SAP |
| **Navigation** | `start_transaction()` | Open transaction |
| | `end_transaction()` | Close transaction |
| | `send_command()` | Execute SAP commands |
| **Discovery** | `get_sap_gui_tree()` | Find element IDs |
| | `find_by_id()` | Verify element exists |
| | `find_by_name()` | Find by name/type |
| **Input** | `set_text()` | Enter text |
| | `set_checkbox()` | Check/uncheck |
| | `set_radio_button()` | Select radio |
| | `set_combobox()` | Select dropdown |
| | `press_button()` | Click button |
| **Output** | `get_text()` | Read field value |
| | `get_grid_data()` | Extract grid data |
| **Grid** | `select_grid_row()` | Select row |
| | `double_click_grid_cell()` | Open details |
| | `set_vertical_scrollbar_position()` | Scroll grid |
| **Window** | `maximize_window()` | Maximize window |
| | `set_focus()` | Focus element |

---

**API Version**: 0.1.0
**Last Updated**: 2025-11-29
**Total MCP Tools**: 28 (all exposed)
**Recent Additions**: login_to_sap() (auto-login capability), get_horizontal_scrollbar_position() (decorator fixed)
