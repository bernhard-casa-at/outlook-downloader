#!/usr/bin/env python3
"""
Diagnose search results and compare different search methods
"""

import sys
from pathlib import Path
import msal
import requests


def load_env_file(env_path: Path) -> dict:
    """Load environment variables from .env file"""
    env_vars = {}

    if not env_path.exists():
        print(f"Error: {env_path} file not found")
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

                if value.startswith('"') and value.endswith('"'):
                    value = value[1:-1]
                elif value.startswith("'") and value.endswith("'"):
                    value = value[1:-1]

                env_vars[key] = value

    return env_vars


def authenticate(tenant_id, client_id, client_secret):
    """Authenticate and get access token"""
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret
    )

    scopes = ["https://graph.microsoft.com/.default"]
    result = app.acquire_token_for_client(scopes=scopes)

    if "access_token" in result:
        return result["access_token"]
    else:
        print(f"Authentication failed: {result.get('error_description', 'Unknown error')}")
        sys.exit(1)


def count_with_search(token, mailbox, search_query):
    """Count emails using $search parameter"""
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/messages"
    headers = {
        'Authorization': f'Bearer {token}',
        'ConsistencyLevel': 'eventual'
    }
    params = {
        '$search': f'"{search_query}"',
        '$count': 'true',
        '$top': 1,
        '$select': 'id'
    }

    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()
    data = response.json()

    # The @odata.count gives us the total count with $count=true
    total_count = data.get('@odata.count')
    return total_count


def count_with_filter(token, mailbox, terms):
    """Count emails using $filter parameter for each term"""
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/messages"
    headers = {'Authorization': f'Bearer {token}'}

    counts = {}

    for term in terms:
        # Try filtering on subject and body
        filter_query = f"contains(subject,'{term}') or contains(body/content,'{term}')"
        params = {
            '$filter': filter_query,
            '$top': 999,
            '$select': 'id'
        }

        all_ids = set()
        page_url = url

        while page_url:
            if page_url == url:
                response = requests.get(page_url, headers=headers, params=params)
            else:
                response = requests.get(page_url, headers=headers)

            response.raise_for_status()
            data = response.json()

            for msg in data.get('value', []):
                all_ids.add(msg['id'])

            page_url = data.get('@odata.nextLink')

        counts[term] = len(all_ids)

    return counts


def list_all_folders(token, mailbox):
    """List all mail folders in the mailbox"""
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders"
    headers = {'Authorization': f'Bearer {token}'}

    response = requests.get(url, headers=headers)
    response.raise_for_status()
    data = response.json()

    folders = []
    for folder in data.get('value', []):
        folders.append({
            'name': folder.get('displayName'),
            'id': folder.get('id'),
            'totalItemCount': folder.get('totalItemCount', 0),
            'unreadItemCount': folder.get('unreadItemCount', 0)
        })

    return folders


def search_all_folders(token, mailbox, search_query):
    """Search in all folders and return counts per folder"""
    folders = list_all_folders(token, mailbox)
    headers = {
        'Authorization': f'Bearer {token}',
        'ConsistencyLevel': 'eventual'
    }

    folder_counts = {}

    for folder in folders:
        folder_name = folder['name']
        folder_id = folder['id']

        url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/{folder_id}/messages"
        params = {
            '$search': f'"{search_query}"',
            '$top': 1,
            '$select': 'id'
        }

        try:
            count = 0
            page_url = url

            while page_url:
                if page_url == url:
                    response = requests.get(page_url, headers=headers, params=params)
                else:
                    response = requests.get(page_url, headers=headers)

                response.raise_for_status()
                data = response.json()

                count += len(data.get('value', []))
                page_url = data.get('@odata.nextLink')

            if count > 0:
                folder_counts[folder_name] = count

        except Exception as e:
            print(f"Error searching folder {folder_name}: {str(e)}")

    return folder_counts


def main():
    print("=" * 70)
    print("Search Diagnostics - Finding all matching emails")
    print("=" * 70)
    print()

    # Load credentials
    env_vars = load_env_file(Path('.env'))
    tenant_id = env_vars['TENANT_ID']
    client_id = env_vars['CLIENT_ID']
    client_secret = env_vars['CLIENT_SECRET']
    mailbox = env_vars.get('MAILBOX_EMAIL', 'emma@emmaandlorenzo.com')
    search_query = env_vars.get('SEARCH_QUERY', 'riproperty.co.uk OR jafri')

    print(f"Mailbox: {mailbox}")
    print(f"Search query: {search_query}")
    print()

    # Authenticate
    print("Authenticating...")
    token = authenticate(tenant_id, client_id, client_secret)
    print("✓ Authenticated")
    print()

    # Method 1: Using $search with $count
    print("Method 1: Using Graph API $search parameter")
    try:
        count = count_with_search(token, mailbox, search_query)
        if count is not None:
            print(f"  Result: {count} emails found")
        else:
            print("  Result: Count not available (will need to paginate through all results)")
    except Exception as e:
        print(f"  Error: {e}")
    print()

    # Method 2: List all folders
    print("Method 2: Checking all mailbox folders")
    try:
        folders = list_all_folders(token, mailbox)
        print(f"  Found {len(folders)} folders:")
        for folder in folders:
            print(f"    - {folder['name']}: {folder['totalItemCount']} total items")
    except Exception as e:
        print(f"  Error: {e}")
    print()

    # Method 3: Search in each folder
    print("Method 3: Searching each folder individually")
    print("  (This may take a minute...)")
    try:
        folder_counts = search_all_folders(token, mailbox, search_query)
        total = sum(folder_counts.values())
        print(f"\n  Results by folder:")
        for folder_name, count in sorted(folder_counts.items(), key=lambda x: x[1], reverse=True):
            print(f"    - {folder_name}: {count} emails")
        print(f"\n  Total across all folders: {total} emails")
    except Exception as e:
        print(f"  Error: {e}")
    print()

    # Method 4: Try individual search terms
    print("Method 4: Searching for individual terms")
    terms = ['riproperty.co.uk', 'jafri']
    try:
        term_counts = count_with_filter(token, mailbox, terms)
        print("  Results by term:")
        for term, count in term_counts.items():
            print(f"    - '{term}': {count} emails")
        print()
        print("  Note: These counts may overlap (same email matching both terms)")
    except Exception as e:
        print(f"  Error: {e}")
    print()

    print("=" * 70)
    print("Diagnosis complete!")
    print()
    print("If the folder search shows more emails than the main search,")
    print("you may need to search specific folders like Archive or Deleted Items.")
    print("=" * 70)


if __name__ == '__main__':
    main()
