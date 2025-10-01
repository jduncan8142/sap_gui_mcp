from setuptools import setup, find_packages

setup(
    name="sap_gui_mcp",
    version="0.1.0",
    description="MCP server for SAP GUI that uses FastMCP and SAP GUI Scripting API",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    python_requires=">=3.13",
)
