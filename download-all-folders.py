#!/usr/bin/env python3
"""
Enhanced Outlook Downloader - Searches ALL folders including nested folders and archive
This version is more comprehensive than the standard search to match compliance search results
"""

import sys
from pathlib import Path
import msal
import requests
import time
import re
import logging
from typing import List, Dict, Set
import base64

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


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
    logger.info("Authenticating with Microsoft Graph API...")

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret
    )

    scopes = ["https://graph.microsoft.com/.default"]
    result = app.acquire_token_for_client(scopes=scopes)

    if "access_token" in result:
        logger.info("Authentication successful")
        return result["access_token"]
    else:
        logger.error(f"Authentication failed: {result.get('error_description', 'Unknown error')}")
        sys.exit(1)


def get_all_folders_recursive(token: str, mailbox: str, parent_id: str = None) -> List[Dict]:
    """Recursively get all folders in the mailbox"""
    headers = {'Authorization': f'Bearer {token}'}

    if parent_id:
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/{parent_id}/childFolders"
    else:
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders"

    all_folders = []

    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        data = response.json()

        for folder in data.get('value', []):
            folder_info = {
                'id': folder['id'],
                'displayName': folder['displayName'],
                'totalItemCount': folder.get('totalItemCount', 0),
                'parentFolderId': folder.get('parentFolderId')
            }
            all_folders.append(folder_info)

            # Recursively get child folders
            child_folders = get_all_folders_recursive(token, mailbox, folder['id'])
            all_folders.extend(child_folders)

    except Exception as e:
        logger.warning(f"Error getting folders for parent {parent_id}: {e}")

    return all_folders


def search_folder(token: str, mailbox: str, folder_id: str, search_query: str) -> List[Dict]:
    """Search for emails in a specific folder"""
    headers = {
        'Authorization': f'Bearer {token}',
        'ConsistencyLevel': 'eventual'
    }

    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/{folder_id}/messages"
    params = {
        '$search': f'"{search_query}"',
        '$top': 50,
        '$select': 'id,subject,from,receivedDateTime,hasAttachments,internetMessageId'
    }

    messages = []
    page_count = 0

    try:
        while url:
            page_count += 1
            response = requests.get(url, headers=headers, params=params if page_count == 1 else None)
            response.raise_for_status()

            data = response.json()
            batch = data.get('value', [])
            messages.extend(batch)

            url = data.get('@odata.nextLink')
            if url:
                time.sleep(0.3)

    except Exception as e:
        logger.warning(f"Error searching folder {folder_id}: {e}")

    return messages


def download_email_as_eml(token: str, mailbox: str, message_id: str, output_path: Path) -> bool:
    """Download email in EML format"""
    try:
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/messages/{message_id}/$value"
        headers = {'Authorization': f'Bearer {token}'}

        response = requests.get(url, headers=headers)
        response.raise_for_status()

        with open(output_path, 'wb') as f:
            f.write(response.content)

        return True

    except Exception as e:
        logger.error(f"Error downloading email {message_id}: {e}")
        return False


def download_attachments(token: str, mailbox: str, message_id: str, output_dir: Path) -> List[str]:
    """Download all attachments for a message"""
    try:
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/messages/{message_id}/attachments"
        headers = {'Authorization': f'Bearer {token}'}

        response = requests.get(url, headers=headers)
        response.raise_for_status()

        attachments = response.json().get('value', [])
        downloaded_files = []

        for attachment in attachments:
            attachment_name = attachment.get('name', 'unnamed_attachment')

            if attachment.get('@odata.type') == '#microsoft.graph.fileAttachment':
                content_bytes = attachment.get('contentBytes')

                if content_bytes:
                    file_path = output_dir / attachment_name
                    counter = 1
                    while file_path.exists():
                        name_parts = attachment_name.rsplit('.', 1)
                        if len(name_parts) == 2:
                            file_path = output_dir / f"{name_parts[0]}_{counter}.{name_parts[1]}"
                        else:
                            file_path = output_dir / f"{attachment_name}_{counter}"
                        counter += 1

                    with open(file_path, 'wb') as f:
                        f.write(base64.b64decode(content_bytes))

                    downloaded_files.append(file_path.name)

        return downloaded_files

    except Exception as e:
        logger.error(f"Error downloading attachments for message {message_id}: {e}")
        return []


def sanitize_filename(filename: str, max_length: int = 200) -> str:
    """Sanitize a string to be used as a filename"""
    sanitized = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', filename)
    sanitized = sanitized.strip('. ')
    if len(sanitized) > max_length:
        sanitized = sanitized[:max_length]
    return sanitized if sanitized else 'unnamed'


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description='Download emails from ALL folders (including nested) in Microsoft 365 mailbox'
    )
    parser.add_argument('--output-dir', default='./emails_all_folders',
                       help='Directory to save emails (default: ./emails_all_folders)')
    parser.add_argument('--attachments-dir', default='./attachments_all_folders',
                       help='Directory to save attachments (default: ./attachments_all_folders)')
    parser.add_argument('--verbose', '-v', action='store_true',
                       help='Enable verbose logging')

    args = parser.parse_args()

    if args.verbose:
        logger.setLevel(logging.DEBUG)

    # Load credentials
    env_vars = load_env_file(Path('.env'))
    tenant_id = env_vars['TENANT_ID']
    client_id = env_vars['CLIENT_ID']
    client_secret = env_vars['CLIENT_SECRET']
    mailbox = env_vars.get('MAILBOX_EMAIL', 'emma@emmaandlorenzo.com')
    search_query = env_vars.get('SEARCH_QUERY', 'riproperty.co.uk OR jafri')

    logger.info(f"Mailbox: {mailbox}")
    logger.info(f"Search query: {search_query}")
    logger.info("")

    # Authenticate
    token = authenticate(tenant_id, client_id, client_secret)

    # Create output directories
    output_dir = Path(args.output_dir)
    attachments_dir = Path(args.attachments_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    attachments_dir.mkdir(parents=True, exist_ok=True)

    # Get all folders recursively
    logger.info("Discovering all folders (including nested)...")
    all_folders = get_all_folders_recursive(token, mailbox)
    logger.info(f"Found {len(all_folders)} total folders")
    logger.info("")

    # Track unique messages by internetMessageId to avoid duplicates
    seen_message_ids: Set[str] = set()
    all_messages = []

    # Search each folder
    for idx, folder in enumerate(all_folders, 1):
        folder_name = folder['displayName']
        folder_id = folder['id']
        item_count = folder['totalItemCount']

        logger.info(f"Searching folder {idx}/{len(all_folders)}: {folder_name} ({item_count} items)")

        messages = search_folder(token, mailbox, folder_id, search_query)

        # Deduplicate by internetMessageId
        new_messages = 0
        for msg in messages:
            msg_id = msg.get('internetMessageId', msg['id'])
            if msg_id not in seen_message_ids:
                seen_message_ids.add(msg_id)
                msg['_folderName'] = folder_name
                all_messages.append(msg)
                new_messages += 1

        if new_messages > 0:
            logger.info(f"  Found {new_messages} new matching emails")

        time.sleep(0.2)

    logger.info("")
    logger.info(f"Total unique emails found: {len(all_messages)}")
    logger.info("")

    # Download all emails
    logger.info("Downloading emails...")
    success_count = 0

    for idx, message in enumerate(all_messages, 1):
        message_id = message.get('id')
        subject = message.get('subject', 'No Subject')
        received_date = message.get('receivedDateTime', '')
        has_attachments = message.get('hasAttachments', False)
        folder_name = message.get('_folderName', 'Unknown')

        logger.info(f"Processing {idx}/{len(all_messages)}: [{folder_name}] {subject[:50]}...")

        # Create filename
        date_prefix = received_date[:10] if received_date else 'unknown_date'
        safe_subject = sanitize_filename(subject, max_length=100)
        safe_folder = sanitize_filename(folder_name, max_length=50)
        eml_filename = f"{date_prefix}_{idx:04d}_{safe_folder}_{safe_subject}.eml"
        eml_path = output_dir / eml_filename

        # Download email
        if download_email_as_eml(token, mailbox, message_id, eml_path):
            success_count += 1

            # Download attachments
            if has_attachments:
                email_attachments_dir = attachments_dir / f"{date_prefix}_{idx:04d}_{safe_subject}"
                email_attachments_dir.mkdir(parents=True, exist_ok=True)

                attachment_files = download_attachments(token, mailbox, message_id, email_attachments_dir)
                if attachment_files:
                    logger.info(f"  Downloaded {len(attachment_files)} attachment(s)")

        time.sleep(0.3)

    logger.info("")
    logger.info(f"Download complete: {success_count}/{len(all_messages)} emails downloaded successfully")
    logger.info(f"Emails saved to: {output_dir}")
    logger.info(f"Attachments saved to: {attachments_dir}")


if __name__ == '__main__':
    main()
