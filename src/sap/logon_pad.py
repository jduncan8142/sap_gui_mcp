import os
import subprocess
import time
from typing import Optional
import win32com.client
import ctypes
from logging import getLogger

logger = getLogger("sap_controller")
MAX_RETRIES = 10
WAIT_TIME = 3.0
POLLING_INTERVAL = 0.5


def get_system_language() -> str | None:
    """
    Get the system's default language code.

    Returns:
        Language code as a string (e.g., "en_US") or None if it cannot be determined
    """
    # Get the system default locale
    kernel32 = ctypes.windll.kernel32
    language_id = kernel32.GetUserDefaultUILanguage()
    system_language_id = kernel32.GetSystemDefaultUILanguage()

    # Convert LANGID to language string
    def langid_to_locale(langid):
        from locale import windows_locale

        return windows_locale.get(langid, None)

    return langid_to_locale(language_id) if langid_to_locale(language_id) else langid_to_locale(system_language_id)


def is_sap_logon_running() -> bool:
    """
    Check if SAP Logon (saplogon.exe) is currently running.

    Returns:
        True if SAP Logon is running, False otherwise
    """
    try:
        # Try to get SAP GUI object - if it exists, SAP Logon is running
        sap_gui = win32com.client.GetObject("SAPGUI")
        return sap_gui is not None
    except Exception:
        return False


def find_sap_logon_path() -> str | None:
    """
    Find the SAP Logon executable path.

    Returns:
        Path to saplogon.exe if found, None otherwise
    """
    # Common SAP Logon installation paths
    common_paths = [
        r"C:\Program Files (x86)\SAP\FrontEnd\SAPGUI\saplogon.exe",
        r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",
        r"C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe",
    ]

    # Check environment variable for custom SAP GUI path
    sap_gui_path = os.getenv("SAP_LOGON_PATH")
    if sap_gui_path and os.path.exists(sap_gui_path):
        return sap_gui_path

    # Check common installation paths
    for path in common_paths:
        if os.path.exists(path):
            return path

    # Try to find in PATH environment variable
    try:
        result = subprocess.run(["where", "saplogon.exe"], capture_output=True, text=True, timeout=5, shell=True)
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip().split("\n")[0]
    except Exception:
        pass

    return None


def launch_sap_logon(wait_time: float = WAIT_TIME, max_retries: int = MAX_RETRIES) -> tuple[bool, str]:
    """
    Launch SAP Logon pad if it's not already running.

    Args:
        wait_time: Time to wait (in seconds) after launching for SAP Logon to initialize

    Returns:
        Tuple of (success: bool, message: str)
    """
    try:
        # Check if SAP Logon is already running
        if is_sap_logon_running():
            return True, "SAP Logon is already running"

        # Find SAP Logon executable
        sap_logon_path = find_sap_logon_path()
        if not sap_logon_path:
            error_msg = (
                "SAP Logon executable not found. Please install SAP GUI or set SAP_LOGON_PATH environment variable."
            )
            logger.error(error_msg)
            return False, error_msg

        # Launch SAP Logon
        logger.info(f"Launching SAP Logon from: {sap_logon_path}")
        subprocess.Popen([sap_logon_path], shell=True)

        # Wait for SAP Logon to initialize
        time.sleep(wait_time)

        # Verify SAP Logon started successfully
        for attempt in range(max_retries):
            if is_sap_logon_running():
                success_msg = f"SAP Logon launched successfully from: {sap_logon_path}"
                logger.info(success_msg)
                return True, success_msg
            time.sleep(POLLING_INTERVAL)

        # SAP Logon didn't start within expected time
        error_msg = "SAP Logon was launched but failed to initialize within expected time"
        logger.warning(error_msg)
        return False, error_msg

    except Exception as e:
        error_msg = f"Failed to launch SAP Logon: {str(e)}"
        logger.error(error_msg)
        return False, error_msg


def create_sap_session(
    system: str,
    client: Optional[str] = None,
    user: Optional[str] = None,
    password: Optional[str] = None,
    language: Optional[str] = None,
    max_wait_time: float = WAIT_TIME,
    use_sso: bool = False,
) -> Optional[win32com.client.CDispatch]:
    """
    Create a new SAP session by logging in with credentials or SSO.

    Args:
        system: SAP system ID (e.g., "PRD", "DEV")
        client: SAP client number (e.g., "100"). Optional if use_sso=True
        user: SAP username. Optional if use_sso=True
        password: SAP password. Optional if use_sso=True
        language: SAP language code (default: "EN")
        max_wait_time: Maximum time to wait for connection in seconds (default: WAIT_TIME=3.0)
        use_sso: If True, use Windows SSO authentication instead of credentials (default: False)

    Returns:
        SAP session object if successful, None otherwise
    """
    try:
        # Get SAP GUI Scripting object
        sap_gui = win32com.client.GetObject("SAPGUI")
        if not sap_gui:
            logger.error("SAP GUI is not running. Please start SAP GUI first.")
            return None

        sap_application = sap_gui.GetScriptingEngine
        if not sap_application:
            logger.error("Failed to get SAP Scripting Engine.")
            return None

        # Build connection string
        # Format: /H/<host>/S/<service> or system description
        # For SAP Logon pad connections, we use the system ID directly
        connection_string = system

        # Open connection
        logger.info(f"Attempting to connect to SAP system: {system}")
        connection = sap_application.OpenConnection(connection_string, True)

        if not connection:
            logger.error(f"Failed to open connection to system: {system}")
            return None

        # Wait for connection to be established
        start_time = time.time()
        while not connection.Sessions and (time.time() - start_time) < max_wait_time:
            time.sleep(POLLING_INTERVAL)

        if not connection.Sessions:
            logger.error("Connection established but no session created.")
            return None

        session = connection.Sessions[0]

        # Handle login based on authentication method
        try:
            # Wait for login window to appear
            time.sleep(max_wait_time)

            if use_sso:
                # SSO Authentication - use Windows credentials
                logger.info(f"Attempting SSO login to {system}")

                # For SSO, we typically only need to set client and press Enter
                # SAP GUI will handle Windows authentication automatically
                if client:
                    session.FindById("wnd[0]/usr/txtRSYST-MANDT").Text = client

                # Press Enter to trigger SSO authentication
                session.FindById("wnd[0]").SendVKey(0)  # VKey 0 = Enter

            else:
                # Standard credential-based authentication
                logger.info(f"Attempting credential login to {system} as {user}")

                # Find and fill login fields
                # Standard SAP login screen field IDs
                session.FindById("wnd[0]/usr/txtRSYST-MANDT").Text = client
                session.FindById("wnd[0]/usr/txtRSYST-BNAME").Text = user
                session.FindById("wnd[0]/usr/pwdRSYST-BCODE").Text = password
                session.FindById("wnd[0]/usr/txtRSYST-LANGU").Text = language

                # Press Enter to login
                session.FindById("wnd[0]").SendVKey(0)  # VKey 0 = Enter

            # Wait for login to complete
            time.sleep(max_wait_time)

            # Check if login was successful by verifying we're not still on login screen
            try:
                # If we can still find the password field, login failed
                session.FindById("wnd[0]/usr/pwdRSYST-BCODE")
                auth_method = "SSO" if use_sso else "credential"
                logger.error(f"Login failed - still on login screen. Check {auth_method} configuration.")
                return None
            except Exception:
                # Password field not found = we've moved past login screen = success
                auth_method = "SSO" if use_sso else f"credentials for {user}"
                logger.info(f"Successfully logged in to {system} using {auth_method}")
                return session

        except Exception as login_error:
            logger.error(f"Error during login process: {str(login_error)}")
            return None

    except Exception as e:
        logger.error(f"Failed to create SAP session: {str(e)}")
        return None


def sap_session(auto_login: bool = False) -> Optional[win32com.client.CDispatch]:
    """
    Get the current SAP session, optionally creating one if it doesn't exist.

    Args:
        auto_login: If True, attempt to create a new session using environment credentials when no existing session is found (default: False)

    Returns:
        SAP session object if available or created, None otherwise
    """
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        if not sap_gui:
            logger.error("SAP GUI is not running.")
            return None

        sap_application = sap_gui.GetScriptingEngine
        if not sap_application:
            logger.error("Failed to get SAP Scripting Engine.")
            return None

        # Check for existing connections and sessions
        if sap_application.Connections and sap_application.Connections.Count > 0:
            connection = sap_application.Connections[0]
            if connection.Sessions and connection.Sessions.Count > 0:
                # Existing session found
                return connection.Sessions[0]

        # No existing session found
        if not auto_login:
            logger.error("No SAP connections or sessions found.")
            return None

        # Attempt auto-login using environment credentials
        logger.info("No existing session found. Attempting auto-login...")

        # Get credentials from environment variables
        sap_system = os.getenv("SAP_SYSTEM")
        sap_client = os.getenv("SAP_CLIENT")
        sap_user = os.getenv("SAP_USER")
        sap_password = os.getenv("SAP_PASSWORD")
        sap_language = os.getenv("SAP_LANGUAGE", "EN")

        # Validate credentials are available
        if not all([sap_system, sap_client, sap_user, sap_password]):
            missing = []
            if not sap_system:
                missing.append("SAP_SYSTEM")
            if not sap_client:
                missing.append("SAP_CLIENT")
            if not sap_user:
                missing.append("SAP_USER")
            if not sap_password:
                missing.append("SAP_PASSWORD")

            logger.error(f"Auto-login failed: Missing required environment variables: {', '.join(missing)}")
            return None

        # Type narrowing: Ensure sap_system is not None before calling create_sap_session
        # This should never happen due to validation above, but satisfies the type checker
        if sap_system is None:
            error_msg = "Login failed: System parameter is required"
            logger.error(error_msg)
            return None

        # Create new session
        return create_sap_session(
            system=sap_system,
            client=sap_client,
            user=sap_user,
            password=sap_password,
            language=sap_language,
        )

    except Exception as e:
        logger.error(f"Error in sap_session: {str(e)}")
        return None
