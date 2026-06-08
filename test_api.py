#!/usr/bin/env python3
"""
Test script for the Rate Sheet Generator API.
Run this to verify the API works correctly.
"""

import json
import requests
from time import sleep

BASE_URL = "http://localhost:8000"

# Test data
SAMPLE_REQUEST = {
    "currency": "CAD",
    "output_type": "all_in",
    "data": {
        "currency": "CAD",
        "institutions": [
            {
                "issuer": "Royal Bank of Canada",
                "available": "available",
                "rating": "R-1 (High) – CDIC",
                "1 Year Fixed": "3.75%",
                "2 Year Fixed": "3.80%",
                "take fi money": "yes"
            },
            {
                "issuer": "TD Bank",
                "available": "available",
                "rating": "R-1 (High) – CDIC",
                "1 Year Fixed": "3.70%",
                "2 Year Fixed": "3.75%",
                "take fi money": "yes"
            },
            {
                "issuer": "EQ Bank",
                "available": "available",
                "rating": "BBB (High) – CDIC",
                "1 Year Fixed": "4.05%",
                "2 Year Fixed": "4.10%",
                "take fi money": "no"
            }
        ]
    }
}

def test_health():
    """Test health endpoint."""
    print("Testing /health endpoint...")
    try:
        response = requests.get(f"{BASE_URL}/health")
        print(f"✓ Health check: {response.json()}")
        return True
    except Exception as e:
        print(f"✗ Health check failed: {e}")
        return False

def test_generate_all_in():
    """Test all_in rate sheet generation."""
    print("\nTesting /generate-rate-sheet (all_in)...")
    try:
        response = requests.post(f"{BASE_URL}/generate-rate-sheet", json=SAMPLE_REQUEST)
        data = response.json()
        if data["success"]:
            print(f"✓ Generated rate sheet: {len(data['html'])} chars")
            print(f"  Message: {data['message']}")
            # Print first 200 chars of HTML
            print(f"  HTML preview: {data['html'][:200]}...")
        else:
            print(f"✗ Failed: {data['error']}")
        return data["success"]
    except Exception as e:
        print(f"✗ Request failed: {e}")
        return False

def test_generate_credit_only():
    """Test credit_only rate sheet generation."""
    print("\nTesting /generate-rate-sheet (credit_only)...")
    try:
        request_data = SAMPLE_REQUEST.copy()
        request_data["output_type"] = "credit_only"
        response = requests.post(f"{BASE_URL}/generate-rate-sheet", json=request_data)
        data = response.json()
        if data["success"]:
            print(f"✓ Generated credit-only sheet: {len(data['html'])} chars")
        else:
            print(f"✗ Failed: {data['error']}")
        return data["success"]
    except Exception as e:
        print(f"✗ Request failed: {e}")
        return False

def test_generate_usd():
    """Test USD rate sheet generation."""
    print("\nTesting USD rate sheet...")
    usd_request = {
        "currency": "USD",
        "output_type": "all_in",
        "data": {
            "currency": "USD",
            "institutions": [
                {
                    "issuer": "Oaken Financial (US)",
                    "available": "available",
                    "DBRS": "BBB",
                    "S&P": "BBB",
                    "1": "4.25%",
                    "2": "4.35%"
                }
            ]
        }
    }
    try:
        response = requests.post(f"{BASE_URL}/generate-rate-sheet", json=usd_request)
        data = response.json()
        if data["success"]:
            print(f"✓ Generated USD sheet: {len(data['html'])} chars")
        else:
            print(f"✗ Failed: {data['error']}")
        return data["success"]
    except Exception as e:
        print(f"✗ Request failed: {e}")
        return False

def main():
    """Run all tests."""
    print("=" * 60)
    print("Rate Sheet Generator API Test Suite")
    print("=" * 60)
    print(f"Base URL: {BASE_URL}")
    print()

    # Wait a moment for user to see the message
    print("Make sure the API is running:")
    print("  python api.py")
    print()
    input("Press Enter to continue with tests...")

    results = []
    results.append(("Health Check", test_health()))
    results.append(("All-In Rate Sheet", test_generate_all_in()))
    results.append(("Credit-Only Rate Sheet", test_generate_credit_only()))
    results.append(("USD Rate Sheet", test_generate_usd()))

    print("\n" + "=" * 60)
    print("Test Summary:")
    print("=" * 60)
    passed = sum(1 for _, result in results if result)
    total = len(results)
    for test_name, result in results:
        status = "✓ PASS" if result else "✗ FAIL"
        print(f"{test_name}: {status}")
    print()
    print(f"Total: {passed}/{total} tests passed")

    if passed == total:
        print("\n🎉 All tests passed! API is ready for Power Automate.")
    else:
        print(f"\n⚠️  {total - passed} test(s) failed. Check the output above.")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nTests cancelled.")
