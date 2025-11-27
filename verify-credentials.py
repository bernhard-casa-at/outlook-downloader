#!/usr/bin/env python3
"""
Verify Azure AD credentials and test authentication
"""

import sys
from pathlib import Path


def load_env_file(env_path: Path) -> dict:
    """Load environment variables from .env file"""
    env_vars = {}

    if not env_path.exists():
        print(f"Error: {env_path} file not found")
        print("Please create a .env file based on .env.example")
        sys.exit(1)

    with open(env_path, 'r') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#'):
                continue

            if '=' in line:
                key, value = line.split('=', 1)
                key = key.strip()
                value = value.strip()

                # Remove quotes if present
                if value.startswith('"') and value.endswith('"'):
                    value = value[1:-1]
                elif value.startswith("'") and value.endswith("'"):
                    value = value[1:-1]

                env_vars[key] = value

    return env_vars


def verify_credentials():
    """Verify credentials from .env file"""
    print("=" * 60)
    print("Azure AD Credentials Verification")
    print("=" * 60)
    print()

    # Load .env file
    env_file = Path('.env')
    env_vars = load_env_file(env_file)

    # Check and display credentials
    print("1. Checking .env file contents:")
    print()

    tenant_id = env_vars.get('TENANT_ID', '')
    client_id = env_vars.get('CLIENT_ID', '')
    client_secret = env_vars.get('CLIENT_SECRET', '')

    if not tenant_id:
        print("  ❌ TENANT_ID is missing or empty")
    else:
        print(f"  ✓ TENANT_ID: {tenant_id[:8]}...{tenant_id[-4:]} (length: {len(tenant_id)})")

    if not client_id:
        print("  ❌ CLIENT_ID is missing or empty")
    else:
        print(f"  ✓ CLIENT_ID: {client_id[:8]}...{client_id[-4:]} (length: {len(client_id)})")

    if not client_secret:
        print("  ❌ CLIENT_SECRET is missing or empty")
    else:
        print(f"  ✓ CLIENT_SECRET: {'*' * 20} (length: {len(client_secret)})")

    print()

    if not all([tenant_id, client_id, client_secret]):
        print("Error: Some required credentials are missing!")
        print()
        print("Please check your .env file and ensure all values are set correctly.")
        sys.exit(1)

    # Check for common issues
    print("2. Checking for common issues:")
    print()

    issues_found = False

    # Check for extra whitespace
    if tenant_id != tenant_id.strip() or client_id != client_id.strip() or client_secret != client_secret.strip():
        print("  ⚠️  Extra whitespace detected in credentials")
        issues_found = True

    # Check for placeholder values
    placeholders = ['your-tenant-id', 'your-client-id', 'your-client-secret', 'YOUR_TENANT_ID']
    if any(p in [tenant_id, client_id, client_secret] for p in placeholders):
        print("  ❌ Placeholder values detected - you need to replace with actual credentials")
        issues_found = True

    # Check GUID format for tenant_id and client_id (should be UUID format)
    import re
    uuid_pattern = re.compile(r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$', re.IGNORECASE)

    if not uuid_pattern.match(tenant_id):
        print(f"  ⚠️  TENANT_ID doesn't match UUID format (might be incorrect)")
        issues_found = True

    if not uuid_pattern.match(client_id):
        print(f"  ⚠️  CLIENT_ID doesn't match UUID format (might be incorrect)")
        issues_found = True

    if not issues_found:
        print("  ✓ No obvious issues detected")

    print()

    # Try authentication
    print("3. Testing authentication with Microsoft Graph API:")
    print()

    try:
        import msal
        import requests

        authority = f"https://login.microsoftonline.com/{tenant_id}"
        app = msal.ConfidentialClientApplication(
            client_id,
            authority=authority,
            client_credential=client_secret
        )

        scopes = ["https://graph.microsoft.com/.default"]
        result = app.acquire_token_for_client(scopes=scopes)

        if "access_token" in result:
            print("  ✓ Authentication successful!")
            print()
            print("Your credentials are correct and working.")
            print()

            # Test a simple API call
            print("4. Testing Graph API access:")
            print()

            mailbox = env_vars.get('MAILBOX_EMAIL', 'emma@emmaandlorenzo.com')
            headers = {'Authorization': f'Bearer {result["access_token"]}'}

            # Try to get user info
            user_url = f"https://graph.microsoft.com/v1.0/users/{mailbox}"
            user_response = requests.get(user_url, headers=headers)

            if user_response.status_code == 200:
                user_data = user_response.json()
                print(f"  ✓ Successfully accessed mailbox: {user_data.get('mail', mailbox)}")
                print(f"  ✓ Display name: {user_data.get('displayName', 'N/A')}")
            elif user_response.status_code == 403:
                print(f"  ⚠️  Access denied to mailbox: {mailbox}")
                print("     Check that your app has Mail.Read permissions and admin consent")
            else:
                print(f"  ⚠️  API call returned status {user_response.status_code}")
                print(f"     Response: {user_response.text[:200]}")

            print()
            print("=" * 60)
            print("You can now run: ./run-from-env.py --verbose")
            print("=" * 60)
            sys.exit(0)

        else:
            print("  ❌ Authentication failed!")
            print()
            error_desc = result.get('error_description', 'Unknown error')
            error_code = result.get('error', 'Unknown')

            print(f"Error: {error_code}")
            print(f"Description: {error_desc}")
            print()

            # Provide helpful hints based on error
            if "AADSTS7000215" in error_desc:
                print("This error means the client secret is invalid.")
                print()
                print("Common causes:")
                print("  1. You used the Secret ID instead of the Secret Value")
                print("  2. The secret has expired")
                print("  3. The secret was copied incorrectly (partial copy, extra characters)")
                print()
                print("Solution:")
                print("  1. Go to Azure Portal > App Registrations > Your App > Certificates & secrets")
                print("  2. Create a NEW client secret")
                print("  3. Copy the VALUE (not the Secret ID) immediately after creation")
                print("  4. Paste it directly into your .env file as CLIENT_SECRET=<value>")
                print("     (no quotes needed unless the value contains spaces)")
            elif "AADSTS700016" in error_desc:
                print("This error means the Client ID or Tenant ID is incorrect.")
                print()
                print("Solution:")
                print("  1. Verify your Tenant ID in Azure Portal > Azure Active Directory > Overview")
                print("  2. Verify your Client ID in Azure Portal > App Registrations > Your App")

            sys.exit(1)

    except ImportError as e:
        print(f"  ❌ Missing required library: {e}")
        print()
        print("Please ensure you've installed the requirements:")
        print("  pip install -r requirements.txt")
        sys.exit(1)
    except Exception as e:
        print(f"  ❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    verify_credentials()
