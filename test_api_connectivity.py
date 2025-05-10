#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Unit test for SiliconFlow API connectivity.
"""
import os
import requests
import base64
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

API_URL  = os.getenv("SF_API_URL",  "https://api.siliconflow.cn/v1/chat/completions")
API_KEY  = os.getenv("SF_API_KEY")
MODEL    = os.getenv("SF_MODEL", "Qwen/Qwen2.5-VL-72B-Instruct")

HEADERS  = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type":  "application/json"
}

# A very small, 1x1 transparent PNG image, base64 encoded
DUMMY_BASE64_IMAGE = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII="

# Minimal payload for testing
# Some VLMs might require a valid image even for a ping, others might have a health check endpoint.
# We'll try a minimal valid request for the chat/completions endpoint.
TEST_PAYLOAD = {
    "model": MODEL,
    "messages": [
        {
            "role": "system",
            "content": "This is a connectivity test."
        },
        {
            "role": "user",
            "content": [
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/png;base64,{DUMMY_BASE64_IMAGE}",
                        "detail": "low" # Use low detail for test
                    }
                },
                {
                    "type": "text",
                    "text": "This is a test instruction. Please respond with a short confirmation if you are reachable."
                }
            ]
        }
    ],
    "stream": False,
    "temperature": 0.1,
    "max_tokens": 10 # Request a very short response
}

def test_api_connection():
    print("--- SiliconFlow API Connectivity Test ---")
    print(f"Attempting to connect to API_URL: {API_URL}")
    print(f"Using MODEL: {MODEL}")
    if not API_KEY:
        print("❌ Error: SF_API_KEY is not set in the environment or .env file.")
        return
    if not API_URL or API_URL == "https://api.siliconflow.cn/v1": # Check for common misconfiguration
        print(f"❌ Error: SF_API_URL is not set correctly. It should be the full chat completions endpoint (e.g., https://api.siliconflow.cn/v1/chat/completions), not just the base URL.")
        print(f"   Current SF_API_URL from .env: {os.getenv('SF_API_URL')}")
        return


    print("\nSending request...")
    try:
        response = requests.post(API_URL, headers=HEADERS, json=TEST_PAYLOAD, timeout=60)
        response.raise_for_status()  # Raises an HTTPError for bad responses (4XX or 5XX)
        
        print("✅ API Connection Successful!")
        print(f"Status Code: {response.status_code}")
        try:
            response_json = response.json()
            print("Response (first 100 chars):")
            print(str(response_json)[:200] + "..." if len(str(response_json)) > 200 else str(response_json))
            
            # Check for specific success indicators if any known
            if "choices" in response_json and len(response_json["choices"]) > 0:
                print("✅ Received a valid-looking response structure.")
            else:
                print("⚠️ Warning: Response structure might not be as expected, but connection was made.")

        except requests.exceptions.JSONDecodeError:
            print("⚠️ Warning: Successful connection, but response was not valid JSON.")
            print("Response Text (first 200 chars):")
            print(response.text[:200] + "..." if len(response.text) > 200 else response.text)

    except requests.exceptions.HTTPError as e:
        print(f"❌ HTTP Error: {e}")
        print(f"Status Code: {e.response.status_code}")
        print("Response Body:")
        print(e.response.text)
    except requests.exceptions.ConnectionError as e:
        print(f"❌ Connection Error: {e}")
        print("   Please check your network connection and the API_URL.")
    except requests.exceptions.Timeout as e:
        print(f"❌ Timeout Error: {e}")
        print("   The request timed out. The API server might be slow or unreachable.")
    except requests.exceptions.RequestException as e:
        print(f"❌ An unexpected error occurred: {e}")
    finally:
        print("\n--- Test Complete ---")

if __name__ == "__main__":
    test_api_connection()
