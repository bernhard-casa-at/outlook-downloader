#!/bin/bash
# Convenience script to run outlook-downloader with credentials from .env file

# Activate virtual environment
source $(which virtualenvwrapper.sh)
workon outlook-downloader

# Load environment variables from .env file if it exists
if [ -f .env ]; then
    export $(cat .env | grep -v '^#' | xargs)
else
    echo "Error: .env file not found"
    echo "Please create a .env file based on .env.example"
    exit 1
fi

# Check required variables
if [ -z "$TENANT_ID" ] || [ -z "$CLIENT_ID" ] || [ -z "$CLIENT_SECRET" ]; then
    echo "Error: Missing required credentials in .env file"
    echo "Please ensure TENANT_ID, CLIENT_ID, and CLIENT_SECRET are set"
    exit 1
fi

# Set defaults if not provided
MAILBOX_EMAIL=${MAILBOX_EMAIL:-"user@example.com"}
SEARCH_QUERY=${SEARCH_QUERY:-"search terms here"}
MESSAGE_CONTENTS_DIR=${MESSAGE_CONTENTS_DIR:-"./emails"}
ATTACHMENTS_DIR=${ATTACHMENTS_DIR:-"./attachments"}

# Run the downloader
python outlook-downloader.py \
    --account "$MAILBOX_EMAIL" \
    --search "$SEARCH_QUERY" \
    --message-contents "$MESSAGE_CONTENTS_DIR" \
    --attachments-directory "$ATTACHMENTS_DIR" \
    --tenant-id "$TENANT_ID" \
    --client-id "$CLIENT_ID" \
    --client-secret "$CLIENT_SECRET" \
    "$@"
