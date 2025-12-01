"""
Test script for SAP GUI screenshot functionality.

This script tests the screenshot capture feature by:
1. Checking if SAP GUI is running and a session is available
2. Capturing a screenshot with automatic naming
3. Capturing a screenshot with a custom path

Prerequisites:
- SAP GUI must be running
- An active SAP session must be available
"""

import os
import sys

# Fix Windows console encoding for Unicode characters
if sys.platform == "win32":
    import codecs

    sys.stdout = codecs.getwriter("utf-8")(sys.stdout.buffer, "strict")
    sys.stderr = codecs.getwriter("utf-8")(sys.stderr.buffer, "strict")

# Add src directory to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from sap.gui import capture_screenshot
from sap.logon_pad import sap_session


def test_screenshot_capture():
    """Test the screenshot capture functionality."""
    print("=" * 60)
    print("SAP GUI Screenshot Functionality Test")
    print("=" * 60)

    # Test 1: Check if SAP session is available
    print("\n[Test 1] Checking for active SAP session...")
    session = sap_session()
    if not session:
        print("❌ FAILED: No active SAP session found.")
        print("   Please start SAP GUI and log in before running this test.")
        return False

    try:
        session_info = f"User: {session.Info.User}, System: {session.Info.SystemName}"
        print(f"✓ SUCCESS: Active session found - {session_info}")
    except Exception as e:
        print(f"⚠ WARNING: Session found but couldn't retrieve info: {e}")

    # Test 2: Capture screenshot with automatic naming
    print("\n[Test 2] Capturing screenshot with automatic naming...")
    success, message = capture_screenshot()
    if success:
        print(f"✓ SUCCESS: {message}")
    else:
        print(f"❌ FAILED: {message}")
        return False

    # Test 3: Capture screenshot with custom path
    print("\n[Test 3] Capturing screenshot with custom path...")
    custom_path = os.path.join(os.getcwd(), "screenshots", "test_custom_screenshot.png")
    success, message = capture_screenshot(output_path=custom_path)
    if success:
        print(f"✓ SUCCESS: {message}")
        # Verify file exists
        if os.path.exists(custom_path):
            file_size = os.path.getsize(custom_path)
            print(f"  File size: {file_size} bytes")
        else:
            print(f"⚠ WARNING: File not found at {custom_path}")
    else:
        print(f"❌ FAILED: {message}")
        return False

    # Test 4: Capture specific window (wnd[0])
    print("\n[Test 4] Capturing screenshot of specific window (wnd[0])...")
    success, message = capture_screenshot(window_id="wnd[0]")
    if success:
        print(f"✓ SUCCESS: {message}")
    else:
        print(f"❌ FAILED: {message}")
        return False

    print("\n" + "=" * 60)
    print("All tests completed successfully!")
    print("=" * 60)
    return True


if __name__ == "__main__":
    try:
        result = test_screenshot_capture()
        sys.exit(0 if result else 1)
    except Exception as e:
        print(f"\n❌ UNEXPECTED ERROR: {e}")
        import traceback

        traceback.print_exc()
        sys.exit(1)
