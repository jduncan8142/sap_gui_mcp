# SAP GUI MCP Test Cases

## Notes

Below is a list of test cases to execute and test the functionality of the SAP GUI MCP Server.
Run them in order and report the status of each as PASS or FAIL, include a screenshot using the capture_screenshot tool of the results for each step.
If a step fails stop execution of the test cases and include any applicable logs.
Only use the SAP GUI MCP server for these tests, except for verifying results and capturing logs.
For verification, screenshots, and logs you can use any MCP servers required.
When I ask you to run the test cases for the SAP GUI MCP Server these are the test cases you should us.
Treat each case as if a user was entering it in a new chat prompt to you and the SAP GUI MCP Server was the only available tool.
Check back here each time for any new cases that may have been added.
Name the screenshots like TC\<x\>\_\<yyyymmdd\>\_\<hhmmss\>\_\<PASS / FAIL\>.jpeg

## Test Cases

1. Open SAP Logon Pad application, using the SAP GUI MCP Server, if it isn't already open.
2. Open a session for "1.2 ERP - RQ2", using Single Sign On (SSO).
3. Run transaction ZVA05N.
4. For ZVA05N Fill in the selection fields:
   - Sales Document Type: ZSPO
   - Sales Organization: 3000
   - Distribution Channel: 50
   - Division: 50
   - Open Orders: true
5. Execute ZVA05N using F8.
6. Capture the ZVA05N results by taking a screenshot.
7. Return to the main menu.
