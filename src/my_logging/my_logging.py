import logging
import os

# Valid log levels
valid_levels = {"DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"}


# Validate and convert to logging constants
def get_log_level(level_str: str, default: str = "ERROR") -> int:
    """Convert string log level to logging constant."""
    level_str = level_str.upper()
    if level_str not in valid_levels:
        level_str = default
    return getattr(logging, level_str)


# Configure logging from environment variables
def configure_logging() -> None:
    """Configure logging levels from environment variables."""
    # Configure root logger
    root_level = get_log_level(os.getenv("LOG_LEVEL", "ERROR").upper())
    logging.basicConfig(
        level=root_level, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    )


# Configure individual logger levels
server_logger = logging.getLogger("server").setLevel(get_log_level(os.getenv("LOG_LEVEL", "ERROR").upper()))
factmcp_logger = logging.getLogger("fastmcp").setLevel(get_log_level(os.getenv("SERVER_LOG_LEVEL", "ERROR").upper()))
mcp_logger = logging.getLogger("mcp").setLevel(get_log_level(os.getenv("MCP_LOG_LEVEL", "ERROR").upper()))
asyncio_logger = logging.getLogger("asyncio").setLevel(get_log_level(os.getenv("ASYNCIO_LOG_LEVEL", "ERROR").upper()))
sap_controller_logger = logging.getLogger("sap_controller").setLevel(
    get_log_level(os.getenv("SAP_CONTROLLER_LOG_LEVEL", "ERROR").upper())
)
sap_logon_pad = logging.getLogger("sap_logon_pad").setLevel(
    get_log_level(os.getenv("SAP_LOGON_PAD_LOG_LEVEL", "ERROR").upper())
)
