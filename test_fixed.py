#!/usr/bin/env python3
"""
Test script for the fixed template filler service.
Tests the split runs bug fix.
"""

import requests
import json

# Test the fixed endpoint with the problematic placeholders
print("Testing fixed fill endpoint with split runs...")
fill_data = {
    "placeholders": {
        "LOAN_AMOUNT": "$25,650,000",
        "ORIG_EXIT_FEE": "$256,500",
        "DD_UW_FEE": "1.00%",
        "LOAN_TERM": "24 months",
        "PROPERTY_ADDRESS": "89 Montauk Highway, East Moriches, NY",
        "LTC_RATIO": "102.1%",
        "DATE": "December 2024",
        "SPONSOR_NAME": "Test Sponsor LLC",
        "PROPERTY_LOCATION_DESCRIPTION": "East Moriches, Suffolk County, NY",
        "SPONSOR_COMPANY": "Test Development Corp",
        "TARGET_CLOSE_DATE": "January 15, 2025"
    },
    "output_filename": "test_fixed.docx"
}

try:
    # Test locally
    print("\nTesting local endpoint...")
    response = requests.post(
        "http://localhost:8000/fill",
        json=fill_data,
        timeout=30
    )
    
    if response.status_code == 200:
        with open("test_fixed_local.docx", "wb") as f:
            f.write(response.content)
        print("✓ Local test successful! Check test_fixed_local.docx")
        print("  Open the file and verify all placeholders were replaced correctly.")
    else:
        print(f"✗ Local test failed: {response.status_code}")
        print(response.text)
        
except requests.exceptions.ConnectionError:
    print("! Local server not running, skipping local test")

# Also prepare a curl command for production testing
print("\nTo test on production, run:")
print("""
curl -k -X POST https://no40s4o88g804o0ko84g4gc0.app9.anant.systems/fill \\
  -H "Content-Type: application/json" \\
  -d '{
    "placeholders": {
      "LOAN_AMOUNT": "$25,650,000",
      "ORIG_EXIT_FEE": "$256,500",
      "DD_UW_FEE": "1.00%",
      "LOAN_TERM": "24 months",
      "PROPERTY_ADDRESS": "89 Montauk Highway, East Moriches, NY"
    }
  }' \\
  --output test_fixed_production.docx
""")