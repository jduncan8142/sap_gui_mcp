# sap_gui_mcp

MCP server for SAP GUI that uses FastMCP and SAP GUI Scripting API.

## Prerequisites

Before installing this project, ensure you have:

1. **Python 3.13 or higher** - This project requires Python 3.13+
2. **Windows Operating System** - Required for SAP GUI Scripting API
3. **SAP GUI for Windows** - Must be installed and configured
4. **SAP GUI Scripting enabled** - Enable scripting in SAP GUI settings

### Enabling SAP GUI Scripting

1. Open SAP Logon
2. Go to Options → Accessibility & Scripting → Scripting
3. Check "Enable scripting"
4. Optionally check "Notify when a script attaches to SAP GUI"

## Installation

### Option 1: Install from source (Recommended for development)

1. **Clone the repository:**

   ```powershell
   git clone https://github.com/jduncan8142/sap_gui_mcp.git
   cd sap_gui_mcp
   ```

2. **Create a virtual environment:**

   ```powershell
   python -m venv venv
   .\venv\Scripts\Activate.ps1
   ```

3. **Install dependencies:**
   ```powershell
   pip install -e .
   ```

### Option 2: Install using pip (if published)

```powershell
pip install sap-gui-mcp
```

### Option 3: Install using uv (faster alternative)

1. **Install uv if not already installed:**

   ```powershell
   pip install uv
   ```

2. **Install the project:**
   ```powershell
   uv pip install -e .
   ```

### Option 4: Install to Claude Desktop with uv

If you want to use this MCP server with Claude Desktop:

1. **Install uv if not already installed:**

   ```powershell
   pip install uv
   ```

2. **Navigate to the project directory:**

   ```powershell
   cd sap_gui_mcp
   ```

3. **Install to Claude Desktop:**
   ```powershell
   uv run fastmcp install claude-desktop .\src\server.py
   ```

This will automatically configure the MCP server to work with Claude Desktop, allowing you to interact with SAP GUI through Claude's interface.

## Configuration

### Basic Setup

1. **Create environment file (optional but recommended):**
   Create a `.env` file in the project root based on `.env.example`:

   ```bash
   # SAP Credentials (used for auto-login functionality)
   SAP_SYSTEM=your_system_id      # e.g., PRD, DEV, QAS
   SAP_CLIENT=your_client         # e.g., 100, 200
   SAP_USER=your_username
   SAP_PASSWORD=your_password
   SAP_LANGUAGE=EN                # Optional: EN, DE, FR, etc.
   ```

   **Note:** With credentials configured, the MCP server can automatically log in to SAP if no session exists.

2. **Verify SAP GUI connection:**
   - **Option A (Manual):** Start SAP GUI and log into a system before using the MCP server
   - **Option B (Auto-login):** Configure `.env` file with credentials, and the server will automatically create a session when needed

## Running the Server

### As a standalone MCP server:

```powershell
python src/server.py
```

### As a module:

```powershell
python -m src.server
```

## Verification

To verify the installation works correctly:

### Option A: Manual Login
1. Start SAP GUI and log into a system manually
2. Run the MCP server
3. Use `get_session_info()` tool to verify the connection

### Option B: Auto-Login
1. Configure `.env` file with SAP credentials
2. Run the MCP server
3. Use `login_to_sap()` tool to create a new session automatically
4. Verify with `get_session_info()` tool

## Dependencies

This project uses the following key dependencies:

- **fastmcp** (≥2.12.3) - FastMCP framework for building MCP servers
- **python-dotenv** (≥1.1.1) - Environment variable management
- **pywin32** (≥311) - Windows COM interface for SAP GUI scripting
- **setuptools** (≥80.9.0) - Package management

## Troubleshooting

### Common Issues:

1. **"SAP GUI is not running" error:**

   - Ensure SAP GUI is started and you're logged into a system
   - Verify SAP GUI scripting is enabled

2. **COM object errors:**

   - Make sure pywin32 is properly installed
   - Try running `python -m pywin32_postinstall -install` after installing pywin32

3. **Python version issues:**

   - This project requires Python 3.13+
   - Check your Python version with `python --version`

4. **Permission errors:**
   - Run PowerShell as Administrator if needed
   - Ensure SAP GUI scripting permissions are granted
