#!/usr/bin/env python3
"""
Wrapper script to run outlook-downloader with credentials from .env file
Handles special characters in credentials better than bash
"""

import os
import sys
import subprocess
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
            # Skip comments and empty lines
            if not line or line.startswith('#'):
                continue

            # Split on first = sign
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


def main():
    # Load .env file
    env_file = Path('.env')
    env_vars = load_env_file(env_file)

    # Check required variables
    required = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET']
    missing = [var for var in required if var not in env_vars or not env_vars[var]]

    if missing:
        print(f"Error: Missing required credentials in .env file: {', '.join(missing)}")
        print("Please ensure TENANT_ID, CLIENT_ID, and CLIENT_SECRET are set")
        sys.exit(1)

    # Set defaults
    mailbox = env_vars.get('MAILBOX_EMAIL', 'user@example.com')
    search_query = env_vars.get('SEARCH_QUERY', 'search terms here')
    message_dir = env_vars.get('MESSAGE_CONTENTS_DIR', './emails')
    attachments_dir = env_vars.get('ATTACHMENTS_DIR', './attachments')

    # Print loaded values for verification (mask sensitive data)
    print("Loaded configuration:")
    print(f"  Tenant ID: {env_vars['TENANT_ID'][:8]}...")
    print(f"  Client ID: {env_vars['CLIENT_ID'][:8]}...")
    print(f"  Client Secret: {'*' * 20} (length: {len(env_vars['CLIENT_SECRET'])})")
    print(f"  Mailbox: {mailbox}")
    print(f"  Search: {search_query}")
    print()

    # Build command
    cmd = [
        'python',
        'outlook-downloader.py',
        '--account', mailbox,
        '--search', search_query,
        '--message-contents', message_dir,
        '--attachments-directory', attachments_dir,
        '--tenant-id', env_vars['TENANT_ID'],
        '--client-id', env_vars['CLIENT_ID'],
        '--client-secret', env_vars['CLIENT_SECRET']
    ]

    # Add any additional arguments passed to this script
    cmd.extend(sys.argv[1:])

    # Run the command
    try:
        result = subprocess.run(cmd)
        sys.exit(result.returncode)
    except KeyboardInterrupt:
        print("\nOperation cancelled by user")
        sys.exit(1)


if __name__ == '__main__':
    main()
