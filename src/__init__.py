"""SAP GUI MCP Server package."""

from dotenv import load_dotenv

from my_logging.my_logging import configure_logging

# Load environment variables
load_dotenv()

# Apply logging configuration
configure_logging()
