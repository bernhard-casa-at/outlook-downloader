#!/usr/bin/env python3
"""
Outlook Downloader - Download emails and attachments from Microsoft 365 using Graph API
"""

import argparse
import sqlite3
import sys
import logging
from datetime import datetime, timezone
from pathlib import Path
import msal
import requests
from typing import Optional, List, Dict
import time
import re


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class StateDB:
    """SQLite-backed store tracking download and delete status per message."""

    def __init__(self, db_path: Path):
        self.conn = sqlite3.connect(db_path)
        self.conn.execute("PRAGMA journal_mode=WAL")
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS messages (
                message_id       TEXT PRIMARY KEY,
                subject          TEXT,
                received_datetime TEXT,
                downloaded_at    TEXT NOT NULL,
                eml_path         TEXT,
                deleted_from_server INTEGER NOT NULL DEFAULT 0,
                deleted_at       TEXT
            )
        """)
        self.conn.commit()

    def is_downloaded(self, message_id: str) -> bool:
        row = self.conn.execute(
            "SELECT 1 FROM messages WHERE message_id = ?", (message_id,)
        ).fetchone()
        return row is not None

    def is_deleted(self, message_id: str) -> bool:
        row = self.conn.execute(
            "SELECT deleted_from_server FROM messages WHERE message_id = ?",
            (message_id,)
        ).fetchone()
        return row is not None and bool(row[0])

    def record_download(self, message_id: str, subject: str,
                        received_datetime: str, eml_path: str):
        now = datetime.now(timezone.utc).isoformat()
        self.conn.execute(
            """
            INSERT OR IGNORE INTO messages
                (message_id, subject, received_datetime, downloaded_at, eml_path)
            VALUES (?, ?, ?, ?, ?)
            """,
            (message_id, subject, received_datetime, now, eml_path),
        )
        self.conn.commit()

    def record_delete(self, message_id: str):
        now = datetime.now(timezone.utc).isoformat()
        self.conn.execute(
            """
            UPDATE messages
            SET deleted_from_server = 1, deleted_at = ?
            WHERE message_id = ?
            """,
            (now, message_id),
        )
        self.conn.commit()

    def close(self):
        self.conn.close()


class OutlookDownloader:
    """Downloads emails and attachments from Microsoft 365 mailbox using Graph API"""

    GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0'

    def __init__(self, tenant_id: str, client_id: str, client_secret: str, mailbox: str):
        """
        Initialize the Outlook Downloader

        Args:
            tenant_id: Azure AD tenant ID
            client_id: Application (client) ID
            client_secret: Client secret value
            mailbox: Email address of the mailbox to access
        """
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.mailbox = mailbox
        self.access_token = None

    def authenticate(self) -> bool:
        """
        Authenticate using MSAL with client credentials flow

        Returns:
            True if authentication successful, False otherwise
        """
        try:
            logger.info("Authenticating with Microsoft Graph API...")

            authority = f"https://login.microsoftonline.com/{self.tenant_id}"
            app = msal.ConfidentialClientApplication(
                self.client_id,
                authority=authority,
                client_credential=self.client_secret
            )

            # Acquire token for Graph API
            scopes = ["https://graph.microsoft.com/.default"]
            result = app.acquire_token_for_client(scopes=scopes)

            if "access_token" in result:
                self.access_token = result["access_token"]
                logger.info("Authentication successful")
                return True
            else:
                logger.error(f"Authentication failed: {result.get('error_description', 'Unknown error')}")
                return False

        except Exception as e:
            logger.error(f"Authentication error: {str(e)}")
            return False

    def _get_headers(self) -> Dict[str, str]:
        """Get headers for API requests"""
        return {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }

    def get_folder_id(self, folder_name: str) -> Optional[str]:
        """
        Resolve a mail folder display name to its Graph API ID.

        Args:
            folder_name: Display name of the folder (e.g. 'DMARC')

        Returns:
            Folder ID string, or None if not found
        """
        url = f"{self.GRAPH_API_ENDPOINT}/users/{self.mailbox}/mailFolders"
        params = {'$filter': f"displayName eq '{folder_name}'", '$top': 1}
        try:
            response = requests.get(url, headers=self._get_headers(), params=params)
            response.raise_for_status()
            folders = response.json().get('value', [])
            if not folders:
                logger.error(f"Folder '{folder_name}' not found in mailbox")
                return None
            folder_id = folders[0]['id']
            logger.info(f"Resolved folder '{folder_name}' to ID: {folder_id[:20]}...")
            return folder_id
        except Exception as e:
            logger.error(f"Error resolving folder '{folder_name}': {str(e)}")
            return None

    def search_emails(self, subject_filter: str, folder_id: Optional[str] = None) -> List[Dict]:
        """
        Fetch emails whose subject contains subject_filter using Graph API $filter with pagination.

        Args:
            subject_filter: Substring to match against the subject field
            folder_id: Optional folder ID to restrict search to a specific folder

        Returns:
            List of email message objects
        """
        logger.info(f"Searching for emails with subject containing: {subject_filter}")

        all_messages = []
        if folder_id:
            url = f"{self.GRAPH_API_ENDPOINT}/users/{self.mailbox}/mailFolders/{folder_id}/messages"
        else:
            url = f"{self.GRAPH_API_ENDPOINT}/users/{self.mailbox}/messages"

        params = {
            '$filter': (
                f"receivedDateTime ge 1900-01-01T00:00:00Z"
                f" and contains(subject,'{subject_filter}')"
            ),
            '$orderby': 'receivedDateTime asc',
            '$top': 50,
            '$count': 'true',
            '$select': 'id,subject,from,receivedDateTime,hasAttachments',
        }

        headers = self._get_headers()
        headers['ConsistencyLevel'] = 'eventual'  # Required for $filter with $count

        page_count = 0

        try:
            while url:
                page_count += 1
                logger.info(f"Fetching page {page_count}...")

                response = requests.get(url, headers=headers, params=params if page_count == 1 else None)
                response.raise_for_status()

                data = response.json()
                messages = data.get('value', [])
                all_messages.extend(messages)

                logger.info(f"Retrieved {len(messages)} messages (total so far: {len(all_messages)})")

                # Get next page URL
                url = data.get('@odata.nextLink')

                # Small delay to avoid rate limiting
                if url:
                    time.sleep(0.5)

            logger.info(f"Search complete. Found {len(all_messages)} total messages")
            return all_messages

        except requests.exceptions.RequestException as e:
            logger.error(f"Error searching emails: {str(e)}")
            if hasattr(e, 'response') and e.response is not None:
                logger.error(f"Response: {e.response.text}")
            return []

    def delete_email(self, message_id: str) -> bool:
        """
        Permanently delete an email from the mailbox.

        Args:
            message_id: ID of the message to delete

        Returns:
            True if deleted successfully, False otherwise
        """
        try:
            url = f"{self.GRAPH_API_ENDPOINT}/users/{self.mailbox}/messages/{message_id}"
            response = requests.delete(url, headers=self._get_headers())
            response.raise_for_status()
            return True
        except Exception as e:
            logger.error(f"Error deleting message {message_id}: {str(e)}")
            return False

    def download_email_as_eml(self, message_id: str, output_path: Path) -> bool:
        """
        Download email in EML format using MIME content

        Args:
            message_id: ID of the message to download
            output_path: Path where to save the EML file

        Returns:
            True if successful, False otherwise
        """
        try:
            url = f"{self.GRAPH_API_ENDPOINT}/users/{self.mailbox}/messages/{message_id}/$value"
            headers = self._get_headers()

            response = requests.get(url, headers=headers)
            response.raise_for_status()

            # Write MIME content to file
            with open(output_path, 'wb') as f:
                f.write(response.content)

            return True

        except Exception as e:
            logger.error(f"Error downloading email {message_id}: {str(e)}")
            return False

    def download_attachments(self, message_id: str, output_dir: Path) -> List[str]:
        """
        Download all attachments for a given message

        Args:
            message_id: ID of the message
            output_dir: Directory where to save attachments

        Returns:
            List of downloaded attachment filenames
        """
        try:
            url = f"{self.GRAPH_API_ENDPOINT}/users/{self.mailbox}/messages/{message_id}/attachments"
            headers = self._get_headers()

            response = requests.get(url, headers=headers)
            response.raise_for_status()

            attachments = response.json().get('value', [])
            downloaded_files = []

            for attachment in attachments:
                attachment_name = attachment.get('name', 'unnamed_attachment')

                # Handle file attachment (not inline or reference)
                if attachment.get('@odata.type') == '#microsoft.graph.fileAttachment':
                    content_bytes = attachment.get('contentBytes')

                    if content_bytes:
                        # Create unique filename if file already exists
                        file_path = output_dir / attachment_name
                        counter = 1
                        while file_path.exists():
                            name_parts = attachment_name.rsplit('.', 1)
                            if len(name_parts) == 2:
                                file_path = output_dir / f"{name_parts[0]}_{counter}.{name_parts[1]}"
                            else:
                                file_path = output_dir / f"{attachment_name}_{counter}"
                            counter += 1

                        # Decode and write attachment
                        import base64
                        with open(file_path, 'wb') as f:
                            f.write(base64.b64decode(content_bytes))

                        downloaded_files.append(file_path.name)
                        logger.debug(f"Downloaded attachment: {file_path.name}")

            return downloaded_files

        except Exception as e:
            logger.error(f"Error downloading attachments for message {message_id}: {str(e)}")
            return []

    def sanitize_filename(self, filename: str, max_length: int = 200) -> str:
        """
        Sanitize a string to be used as a filename

        Args:
            filename: Original filename
            max_length: Maximum length for filename

        Returns:
            Sanitized filename
        """
        # Remove or replace invalid characters
        sanitized = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', filename)
        # Remove leading/trailing spaces and dots
        sanitized = sanitized.strip('. ')
        # Limit length
        if len(sanitized) > max_length:
            sanitized = sanitized[:max_length]
        return sanitized if sanitized else 'unnamed'

    def process_emails(self, search_query: str, message_contents_dir: Path,
                       attachments_dir: Optional[Path] = None,
                       delete_after_download: bool = False,
                       state_db: Optional['StateDB'] = None,
                       folder_name: Optional[str] = None) -> int:
        """
        Search and download emails with their attachments

        Args:
            search_query: Substring to match against email subjects
            message_contents_dir: Directory to save EML files
            attachments_dir: Directory to save attachments (optional)
            delete_after_download: Delete each email from server after successful download
            state_db: Optional StateDB instance to track progress and prevent duplicates
            folder_name: Optional folder display name to restrict processing (e.g. 'DMARC')

        Returns:
            Number of emails successfully processed
        """
        # Authenticate first
        if not self.authenticate():
            logger.error("Authentication failed. Cannot proceed.")
            return 0

        # Create output directories
        message_contents_dir.mkdir(parents=True, exist_ok=True)
        if attachments_dir:
            attachments_dir.mkdir(parents=True, exist_ok=True)

        # Resolve folder if specified
        folder_id = None
        if folder_name:
            folder_id = self.get_folder_id(folder_name)
            if folder_id is None:
                return 0

        # Search for emails
        messages = self.search_emails(search_query, folder_id=folder_id)

        if not messages:
            logger.warning("No messages found matching the search criteria")
            return 0

        # Process each message
        success_count = 0

        for idx, message in enumerate(messages, 1):
            message_id = message.get('id')
            subject = message.get('subject', 'No Subject')
            received_date = message.get('receivedDateTime', '')
            has_attachments = message.get('hasAttachments', False)

            logger.info(f"Processing {idx}/{len(messages)}: {subject[:50]}...")

            # Skip if already fully processed
            if state_db and state_db.is_deleted(message_id):
                logger.info(f"Skipping (already downloaded and deleted): {subject[:50]}")
                success_count += 1
                continue

            already_downloaded = state_db and state_db.is_downloaded(message_id)

            if not already_downloaded:
                # Create filename from date and subject
                date_prefix = received_date[:10] if received_date else 'unknown_date'
                safe_subject = self.sanitize_filename(subject, max_length=100)
                eml_filename = f"{date_prefix}_{idx:04d}_{safe_subject}.eml"
                eml_path = message_contents_dir / eml_filename

                if not self.download_email_as_eml(message_id, eml_path):
                    logger.warning(f"Failed to download email: {subject[:50]}")
                    time.sleep(0.3)
                    continue

                logger.info(f"Saved email to: {eml_filename}")
                success_count += 1

                if state_db:
                    state_db.record_download(message_id, subject, received_date, str(eml_path))

                # Download attachments if requested and available
                if attachments_dir and has_attachments:
                    email_attachments_dir = attachments_dir / f"{date_prefix}_{idx:04d}_{safe_subject}"
                    email_attachments_dir.mkdir(parents=True, exist_ok=True)
                    attachment_files = self.download_attachments(message_id, email_attachments_dir)
                    if attachment_files:
                        logger.info(f"Downloaded {len(attachment_files)} attachment(s)")
            else:
                logger.info(f"Skipping download (already on disk): {subject[:50]}")
                success_count += 1

            # Delete from server if requested (runs for both fresh and retry cases)
            if delete_after_download:
                if self.delete_email(message_id):
                    logger.info(f"Deleted from server: {subject[:50]}")
                    if state_db:
                        state_db.record_delete(message_id)
                else:
                    logger.warning(f"Failed to delete from server: {subject[:50]}")

            # Small delay to avoid rate limiting
            time.sleep(0.3)

        logger.info(f"\nProcessing complete: {success_count}/{len(messages)} emails downloaded successfully")
        return success_count


def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description='Download emails and attachments from Microsoft 365 mailbox using Graph API',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Example usage:
  %(prog)s --account user@example.com \\
    --search "searchterm1 OR searchterm2" \\
    --message-contents ./emails \\
    --attachments-directory ./attachments \\
    --tenant-id YOUR_TENANT_ID \\
    --client-id YOUR_CLIENT_ID \\
    --client-secret YOUR_CLIENT_SECRET
        """
    )

    parser.add_argument('--account', required=True,
                       help='Email address of the mailbox to access')
    parser.add_argument('--search', required=True,
                       help='Search query string (e.g., "searchterm1 OR searchterm2")')
    parser.add_argument('--message-contents', required=True,
                       help='Directory to save email messages as EML files')
    parser.add_argument('--attachments-directory',
                       help='Directory to save attachments (optional)')

    # Authentication parameters
    parser.add_argument('--tenant-id', required=True,
                       help='Azure AD tenant ID')
    parser.add_argument('--client-id', required=True,
                       help='Application (client) ID')
    parser.add_argument('--client-secret', required=True,
                       help='Client secret value')

    parser.add_argument('--folder', required=True,
                       help='Mailbox folder to process (e.g. "DMARC" or "Inbox")')

    # Optional parameters
    parser.add_argument('--delete-after-download', action='store_true',
                       help='Delete each email from the server after successful download')
    parser.add_argument('--state-db', default='./downloader-state.db',
                       help='Path to SQLite state database (default: ./downloader-state.db)')
    parser.add_argument('--verbose', '-v', action='store_true',
                       help='Enable verbose logging')

    args = parser.parse_args()

    # Set logging level
    if args.verbose:
        logger.setLevel(logging.DEBUG)

    # Convert paths
    message_contents_path = Path(args.message_contents)
    attachments_path = Path(args.attachments_directory) if args.attachments_directory else None

    # Create downloader instance
    downloader = OutlookDownloader(
        tenant_id=args.tenant_id,
        client_id=args.client_id,
        client_secret=args.client_secret,
        mailbox=args.account
    )

    state_db = StateDB(Path(args.state_db))

    # Process emails
    try:
        count = downloader.process_emails(
            search_query=args.search,
            message_contents_dir=message_contents_path,
            attachments_dir=attachments_path,
            delete_after_download=args.delete_after_download,
            state_db=state_db,
            folder_name=args.folder
        )

        if count > 0:
            logger.info(f"Successfully downloaded {count} emails")
            sys.exit(0)
        else:
            logger.error("No emails were downloaded")
            sys.exit(1)

    except KeyboardInterrupt:
        logger.info("\nOperation cancelled by user")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}", exc_info=True)
        sys.exit(1)
    finally:
        state_db.close()


if __name__ == '__main__':
    main()
