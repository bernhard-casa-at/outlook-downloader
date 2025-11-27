#!/usr/bin/env python3
"""
Comprehensive Outlook Downloader - Searches all possible locations
Combines: main messages endpoint + archive mailbox + special folders
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


def search_messages_endpoint(token: str, mailbox: str, search_query: str, endpoint_name: str, url: str) -> List[Dict]:
    """Search using a specific endpoint"""
    logger.info(f"Searching {endpoint_name}...")

    headers = {
        'Authorization': f'Bearer {token}',
        'ConsistencyLevel': 'eventual'
    }

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

            # Tag messages with source
            for msg in batch:
                msg['_source'] = endpoint_name

            messages.extend(batch)
            logger.info(f"  Page {page_count}: {len(batch)} emails (total: {len(messages)})")

            url = data.get('@odata.nextLink')
            if url:
                time.sleep(0.3)

        logger.info(f"  {endpoint_name} total: {len(messages)} emails")
        return messages

    except Exception as e:
        logger.warning(f"Error searching {endpoint_name}: {e}")
        return []


def search_special_folder(token: str, mailbox: str, search_query: str, folder_name: str) -> List[Dict]:
    """Search a special folder by name"""
    logger.info(f"Searching special folder: {folder_name}...")

    headers = {
        'Authorization': f'Bearer {token}',
        'ConsistencyLevel': 'eventual'
    }

    # Try to get the folder by well-known name
    try:
        folder_url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/{folder_name}"
        folder_response = requests.get(folder_url, headers=headers)

        if folder_response.status_code == 200:
            folder_data = folder_response.json()
            folder_id = folder_data['id']

            # Search in this folder
            messages_url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/{folder_id}/messages"
            messages = search_messages_endpoint(token, mailbox, search_query, f"Special folder: {folder_name}", messages_url)
            return messages
        else:
            logger.info(f"  Folder '{folder_name}' not accessible or doesn't exist")
            return []

    except Exception as e:
        logger.warning(f"Error accessing folder {folder_name}: {e}")
        return []


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
        description='Comprehensive email download from all accessible locations'
    )
    parser.add_argument('--output-dir', default='./emails_comprehensive',
                       help='Directory to save emails (default: ./emails_comprehensive)')
    parser.add_argument('--attachments-dir', default='./attachments_comprehensive',
                       help='Directory to save attachments (default: ./attachments_comprehensive)')
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
    mailbox = env_vars.get('MAILBOX_EMAIL', 'user@example.com')
    search_query = env_vars.get('SEARCH_QUERY', 'search terms here')

    logger.info("=" * 70)
    logger.info("COMPREHENSIVE EMAIL DOWNLOAD")
    logger.info("=" * 70)
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

    # Collect messages from all sources
    all_messages = []

    # 1. Search main messages endpoint (this is what got 275 originally)
    main_url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/messages"
    main_messages = search_messages_endpoint(token, mailbox, search_query, "Main mailbox", main_url)
    all_messages.extend(main_messages)
    logger.info("")

    # 2. Try to search archive mailbox
    archive_messages = search_special_folder(token, mailbox, search_query, "archive")
    all_messages.extend(archive_messages)
    logger.info("")

    # 3. Try to search recoverable items (deleted items that can be recovered)
    recoverable_messages = search_special_folder(token, mailbox, search_query, "recoverableitemsdeletions")
    all_messages.extend(recoverable_messages)
    logger.info("")

    # 4. Try other special folders
    special_folders = ['msgfolderroot', 'deleteditems', 'drafts', 'inbox', 'sentitems']
    for folder in special_folders:
        folder_messages = search_special_folder(token, mailbox, search_query, folder)
        all_messages.extend(folder_messages)
        logger.info("")

    # Deduplicate by internetMessageId
    logger.info("Deduplicating messages...")
    seen_ids: Set[str] = set()
    unique_messages = []

    sources_count = {}

    for msg in all_messages:
        msg_id = msg.get('internetMessageId', msg['id'])
        if msg_id not in seen_ids:
            seen_ids.add(msg_id)
            unique_messages.append(msg)

            source = msg.get('_source', 'unknown')
            sources_count[source] = sources_count.get(source, 0) + 1

    logger.info(f"Total messages before deduplication: {len(all_messages)}")
    logger.info(f"Unique messages after deduplication: {len(unique_messages)}")
    logger.info("")
    logger.info("Messages by source:")
    for source, count in sorted(sources_count.items(), key=lambda x: x[1], reverse=True):
        logger.info(f"  {source}: {count}")
    logger.info("")

    # Download all unique emails
    logger.info("Downloading emails...")
    success_count = 0

    for idx, message in enumerate(unique_messages, 1):
        message_id = message.get('id')
        subject = message.get('subject', 'No Subject')
        received_date = message.get('receivedDateTime', '')
        has_attachments = message.get('hasAttachments', False)
        source = message.get('_source', 'unknown')

        logger.info(f"Processing {idx}/{len(unique_messages)}: [{source}] {subject[:50]}...")

        # Create filename
        date_prefix = received_date[:10] if received_date else 'unknown_date'
        safe_subject = sanitize_filename(subject, max_length=100)
        safe_source = sanitize_filename(source, max_length=30)
        eml_filename = f"{date_prefix}_{idx:04d}_{safe_source}_{safe_subject}.eml"
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
    logger.info("=" * 70)
    logger.info(f"DOWNLOAD COMPLETE")
    logger.info("=" * 70)
    logger.info(f"Successfully downloaded: {success_count}/{len(unique_messages)} emails")
    logger.info(f"Emails saved to: {output_dir}")
    logger.info(f"Attachments saved to: {attachments_dir}")
    logger.info("")
    logger.info(f"Original compliance search: 317 emails")
    logger.info(f"This script found: {len(unique_messages)} emails")
    if len(unique_messages) < 317:
        missing = 317 - len(unique_messages)
        logger.info(f"Missing: {missing} emails")
        logger.info("")
        logger.info("The missing emails may be in:")
        logger.info("  - Archive mailbox (separate from main mailbox)")
        logger.info("  - Purged items (permanently deleted but in retention)")
        logger.info("  - Compliance holds (only accessible via compliance tools)")
        logger.info("  - Locations that require additional API permissions")


if __name__ == '__main__':
    main()
