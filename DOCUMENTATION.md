# SAP GUI MCP - Project Documentation

## Project Overview

**SAP GUI MCP** is a Model Context Protocol (MCP) server that enables programmatic interaction with SAP GUI for Windows through the SAP GUI Scripting API. This project bridges AI assistants like Claude with SAP systems, allowing automated interaction with SAP GUI applications.

### Key Information

- **Project Name**: sap_gui_mcp
- **Version**: 0.1.0
- **License**: MIT License (Copyright 2025 Jason)
- **Repository**: https://github.com/jduncan8142/sap_gui_mcp
- **Framework**: FastMCP (≥2.12.3)
- **Python Requirement**: Python 3.13+
- **Platform**: Windows only (requires SAP GUI for Windows)

## Architecture

### Technology Stack

The project is built using:

1. **FastMCP Framework**: Provides the MCP server infrastructure
2. **pywin32**: Enables COM interface access to SAP GUI Scripting API
3. **python-dotenv**: Manages environment configuration
4. **SAP GUI Scripting API**: The underlying COM interface for SAP GUI automation

### Project Structure

```
sap_gui_mcp/
├── src/
│   ├── __init__.py          # Package initialization
│   └── server.py            # Main MCP server implementation
├── .env.example             # Example environment configuration
├── .gitignore              # Git ignore rules
├── .python-version         # Python version specification
├── LICENSE                 # MIT License
├── README.md               # User-facing documentation
├── pyproject.toml          # Project configuration (modern)
├── setup.py                # Package setup (legacy support)
└── uv.lock                 # UV package manager lock file
```

### Core Components

#### 1. Server Module (`src/server.py`)

The main server module contains:

- **FastMCP Server Instance**: `mcp = FastMCP("SAP GUI MCP Server")`
- **Session Management**: `sap_session()` function to retrieve active SAP session
- **Object Tree Parsing**: `sap_object_tree_as_json()` for GUI tree conversion
- **26 MCP Tools**: Decorated with `@mcp.tool()` for client interaction

## Current Features

### Session Management Tools

1. **`get_session_info()`**
   - Returns session metadata (User, Client, Language, System)
   - Useful for verification and debugging

2. **`check_gui_busy()`**
   - Checks if SAP GUI is processing a request
   - Essential for synchronization

### Navigation Tools

3. **`start_transaction(transaction_code: str)`**
   - Launches SAP transactions (e.g., SE38, VA01)
   - Core navigation functionality

4. **`end_transaction()`**
   - Closes the current transaction
   - Clean state management

### Element Discovery Tools

5. **`get_sap_gui_tree()`**
   - Returns JSON representation of the GUI object tree
   - Critical for element identification

6. **`find_by_id(element_id: str, raise_error: bool = False)`**
   - Locates GUI elements by ID
   - Primary element lookup method

7. **`find_by_name(element_name: str, element_type: str)`**
   - Finds elements by name and type
   - Alternative lookup method

8. **`find_all_by_name(element_name: str, element_type: str)`**
   - Returns all matching elements
   - Useful for repeated elements

### Input/Output Tools

9. **`set_text(element_id: str, text: str)`**
   - Sets text field values
   - Basic data entry

10. **`get_text(element_id: str)`**
    - Retrieves text from fields
    - Data extraction

11. **`press_button(element_id: str)`**
    - Simulates button clicks
    - User interaction

12. **`set_radio_button(element_id: str, selected: bool)`**
    - Selects radio buttons
    - Form interaction

13. **`set_checkbox(element_id: str, state: bool)`**
    - Sets checkbox state
    - Boolean input

14. **`set_combobox(element_id: str, key: str)`**
    - Sets dropdown/combobox values
    - Selection input

15. **`set_focus(element_id: str)`**
    - Sets focus to an element
    - Interaction preparation

### Command Execution Tools

16. **`send_command(command: str)`**
    - Sends synchronous commands to SAP
    - Direct SAP command execution

17. **`send_command_async(command: str)`**
    - Sends asynchronous commands
    - Non-blocking operations

### Grid Control Tools

18. **`get_grid_data(element_id: str)`**
    - Extracts all data from grid controls (ALV tables)
    - Returns columns, rows, and cell values
    - Critical for data extraction

19. **`get_grid_cell_value(element_id: str, row: int, column: int)`**
    - Retrieves specific cell value
    - Targeted data access

20. **`select_grid_row(element_id: str, row: int)`**
    - Selects a specific row in grid
    - Row manipulation

21. **`get_selected_grid_rows(element_id: str)`**
    - Returns data from selected rows
    - Batch data extraction

22. **`double_click_grid_cell(element_id: str, row: int, column: int)`**
    - Double-clicks cell to drill down
    - Navigation within grids

### Scrollbar Management Tools

23. **`get_vertical_scrollbar_position(element_id: str)`**
    - Gets vertical scroll position
    - Grid navigation state

24. **`set_vertical_scrollbar_position(element_id: str, position: int)`**
    - Sets vertical scroll position
    - Grid navigation control

25. **`set_horizontal_scrollbar_position(element_id: str, position: int)`**
    - Sets horizontal scroll position
    - Wide grid navigation

26. **`get_horizontal_scrollbar_position(element_id: str)`** (defined but not decorated)
    - Gets horizontal scroll position
    - **Note**: Missing `@mcp.tool()` decorator

### Window Management Tools

27. **`maximize_window()`**
    - Maximizes the active SAP window
    - Screen optimization

## Dependencies

### Required Packages

```toml
dependencies = [
    "fastmcp>=2.12.3",      # MCP server framework
    "python-dotenv>=1.1.1",  # Environment management
    "pywin32>=311",          # Windows COM interface
    "setuptools>=80.9.0",    # Package management
]
```

### Platform Requirements

- **Operating System**: Windows (SAP GUI Scripting API requirement)
- **Python**: 3.13 or higher
- **SAP GUI**: SAP GUI for Windows with scripting enabled

## Installation & Configuration

### Prerequisites Checklist

- [ ] Windows operating system
- [ ] Python 3.13+ installed
- [ ] SAP GUI for Windows installed
- [ ] SAP GUI Scripting enabled in SAP Logon settings

### Installation Methods

#### Method 1: Development Installation
```powershell
git clone https://github.com/jduncan8142/sap_gui_mcp.git
cd sap_gui_mcp
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -e .
```

#### Method 2: UV Package Manager
```powershell
pip install uv
uv pip install -e .
```

#### Method 3: Claude Desktop Integration
```powershell
uv run fastmcp install claude-desktop .\src\server.py
```

### Environment Configuration

The `.env.example` file shows available configuration options:

```env
# SAP Credentials (currently not used in code)
SAP_SYSTEM=your_system
SAP_CLIENT=your_client
SAP_USER=your_username
SAP_PASSWORD=your_password

# Logging Levels
LOG_LEVEL=ERROR
SERVER_LOG_LEVEL=ERROR
MCP_LOG_LEVEL=ERROR
ASYNCIO_LOG_LEVEL=ERROR
SAP_CONTROLLER_LOG_LEVEL=ERROR
```

**Note**: Credential configuration is defined but not currently implemented in the code.

## Usage Patterns

### Basic Workflow

1. **Start SAP GUI and log in** to your system
2. **Launch the MCP server**: `python src/server.py`
3. **Connect MCP client** (e.g., Claude Desktop)
4. **Use tools** to interact with SAP GUI

### Example Use Cases

#### 1. Session Verification
```python
# Check current session
get_session_info()
# Returns: SessionId, User, Client, Language, SystemName, SystemNumber
```

#### 2. Transaction Navigation
```python
# Start a transaction
start_transaction("SE38")

# Work with the transaction
# ...

# End when done
end_transaction()
```

#### 3. Data Entry
```python
# Get GUI tree to find element IDs
get_sap_gui_tree()

# Set field values
set_text("wnd[0]/usr/txtRF02D-KUNNR", "12345")
set_checkbox("wnd[0]/usr/chkPARAM-CHECK", True)
set_combobox("wnd[0]/usr/cmbVARI-MONTH", "01")

# Press button
press_button("wnd[0]/usr/btnEXECUTE")
```

#### 4. Grid Data Extraction
```python
# Extract all data from a grid
grid_data = get_grid_data("wnd[0]/usr/cntlGRID1/shellcont/shell")

# Select specific row
select_grid_row("wnd[0]/usr/cntlGRID1/shellcont/shell", 5)

# Get selected rows data
selected_data = get_selected_grid_rows("wnd[0]/usr/cntlGRID1/shellcont/shell")
```

## Development History

### Recent Commits

```
c09f13f - Removed unused imports
ccdb13f - Added additional tools
31aab4c - Added set_checkbox function
b187ffc - Initial Commit 20250930
62f2ddd - Initial commit
```

### Current Development Branch

- **Active Branch**: `claude/document-sap-gui-mcp-01A3hKVtaAozaP1vSLAqvLV4`
- **Status**: Clean working directory

## Known Issues & Limitations

### Code Issues

1. **Missing Decorator**: `get_horizontal_scrollbar_position()` at line 667 is defined but missing `@mcp.tool()` decorator
2. **Unused Environment Variables**: SAP credentials in `.env.example` are not used in the code
3. **Inconsistent Casing**: Some methods use `findById` while `sap_session()` uses proper casing

### Platform Limitations

1. **Windows Only**: Requires Windows OS due to COM interface dependency
2. **SAP GUI Required**: Must have active SAP GUI session running
3. **Single Session**: Currently connects to first available session (`Connections[0].Sessions[0]`)

### Functional Gaps

1. **No Authentication**: No automatic login capability (credentials unused)
2. **Limited Error Recovery**: Basic error handling without retry logic
3. **No Session Creation**: Assumes existing SAP session
4. **Single Window Focus**: Primarily works with active window

## Security Considerations

### Current Implementation

- Uses Windows COM interface (local system access)
- No authentication implemented in code
- Environment variables prepared but unused
- Local-only operation (no network exposure by default)

### Best Practices

1. **Never commit `.env` files** with real credentials
2. **Use Windows security** to protect SAP GUI access
3. **Audit MCP tool usage** in production environments
4. **Limit MCP server access** to trusted clients

## Testing & Verification

### Manual Testing Steps

1. Start SAP GUI and log into a system
2. Run the MCP server: `python src/server.py`
3. Connect an MCP client
4. Test basic tools:
   - `get_session_info()` - Verify connection
   - `get_sap_gui_tree()` - Check GUI parsing
   - `start_transaction("SE38")` - Test navigation
   - `end_transaction()` - Test cleanup

### Validation Checklist

- [ ] SAP GUI scripting enabled
- [ ] Python 3.13+ available
- [ ] pywin32 properly installed
- [ ] SAP session active before server start
- [ ] MCP server starts without errors
- [ ] Client can connect and list tools

## Future Enhancements

### Recommended Additions

1. **Authentication**: Implement automatic SAP logon using credentials
2. **Multi-Session Support**: Handle multiple SAP connections/sessions
3. **Error Recovery**: Add retry logic and better error messages
4. **Logging Configuration**: Use the defined log levels from environment
5. **Missing Decorator**: Add `@mcp.tool()` to `get_horizontal_scrollbar_position()`
6. **Transaction Automation**: High-level wrappers for common SAP tasks
7. **Screenshot Capability**: Capture SAP GUI screens for documentation
8. **Recording/Playback**: Record user actions for script generation
9. **Validation Tools**: Pre-flight checks for SAP connectivity
10. **Documentation**: Add API documentation with examples for each tool

### Architecture Improvements

1. **Separation of Concerns**: Split server.py into multiple modules
2. **Configuration Management**: Implement proper config loading
3. **Connection Pooling**: Support multiple SAP systems
4. **Async Support**: Better async handling for long-running operations
5. **Type Hints**: More comprehensive type annotations
6. **Unit Tests**: Add test coverage for core functions

## Troubleshooting

### Common Issues

| Issue | Cause | Solution |
|-------|-------|----------|
| "SAP GUI is not running" | SAP not started | Start SAP GUI and log in |
| COM object errors | pywin32 not installed | Run `python -m pywin32_postinstall -install` |
| Python version errors | Python < 3.13 | Upgrade to Python 3.13+ |
| Permission errors | Insufficient rights | Run PowerShell as Administrator |
| Scripting disabled | SAP setting | Enable in SAP Logon options |

### Debug Mode

To enable debug logging, modify `src/server.py`:

```python
logging.basicConfig(level=logging.DEBUG)
```

Or set in `.env`:
```env
LOG_LEVEL=DEBUG
```

## Contributing

### Development Setup

1. Fork the repository
2. Create a feature branch
3. Install in development mode: `pip install -e .`
4. Make changes and test
5. Submit pull request

### Code Style

- Follow PEP 8 guidelines
- Use type hints where applicable
- Add docstrings to all tools
- Log errors and important operations
- Handle exceptions gracefully

## Resources

### Documentation Links

- [FastMCP Documentation](https://github.com/jlowin/fastmcp)
- [SAP GUI Scripting API Guide](https://help.sap.com/doc/saphelp_nw75/7.5.5/en-US/ba/b8bfce3d7e42e49e0bda8de4aa01f4/frameset.htm)
- [pywin32 Documentation](https://github.com/mhammond/pywin32)
- [MCP Protocol Specification](https://modelcontextprotocol.io/)

### Related Projects

- **FastMCP**: The framework powering this server
- **MCP Specification**: The protocol standard
- **SAP GUI Scripting**: The underlying automation API

---

**Last Updated**: 2025-11-29
**Document Version**: 1.0
**Project Status**: Active Development
