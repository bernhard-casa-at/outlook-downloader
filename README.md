# Outlook Downloader

A Python utility to download emails and attachments from Microsoft 365 mailboxes using the Microsoft Graph API.

## Features

- Search emails using Microsoft Graph API search queries
- Download emails in EML format (preserves original MIME content)
- Download attachments with unique filenames
- Handle pagination for large result sets
- Client credentials authentication using MSAL
- Comprehensive error handling and logging

## Prerequisites

- Python 3.11+
- Azure AD application with the following:
  - Tenant ID
  - Client ID (Application ID)
  - Client Secret
  - API Permissions: `Mail.Read` (Application permission)

## Setup

1. **Create virtual environment** (if not already done):
   ```bash
   mkvirtualenv outlook-downloader -p python3.11
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure Azure AD Application**:
   - Register an application in Azure Portal
   - Add API permissions: Microsoft Graph > Application permissions > Mail.Read
   - Grant admin consent for the permissions
   - Create a client secret and note the value

## Usage

### Basic Usage

```bash
python outlook-downloader.py \
  --account emma@emmaandlorenzo.com \
  --search "riproperty.co.uk OR jafri" \
  --message-contents ./emails \
  --attachments-directory ./attachments \
  --tenant-id YOUR_TENANT_ID \
  --client-id YOUR_CLIENT_ID \
  --client-secret YOUR_CLIENT_SECRET
```

### Command-line Arguments

**Required arguments:**
- `--account <email>` - Email address of the mailbox to access
- `--search <query>` - Search query string (supports Boolean operators: AND, OR, NOT)
- `--message-contents <directory>` - Directory to save email messages as EML files
- `--tenant-id <id>` - Azure AD tenant ID
- `--client-id <id>` - Application (client) ID
- `--client-secret <secret>` - Client secret value

**Optional arguments:**
- `--attachments-directory <directory>` - Directory to save attachments
- `--verbose, -v` - Enable verbose logging

### Search Query Examples

```bash
# Search for emails containing specific domain
--search "riproperty.co.uk"

# Search with OR operator
--search "riproperty.co.uk OR jafri"

# Search with AND operator
--search "invoice AND urgent"

# Search in specific fields
--search "from:john@example.com"
--search "subject:meeting"
```

### Output Format

**Email files:**
- Saved as EML files with format: `YYYY-MM-DD_NNNN_Subject.eml`
- NNNN is a 4-digit sequence number
- Subject is sanitized for filesystem compatibility
- EML files can be opened in Outlook, Thunderbird, or any email client

**Attachments:**
- Saved in subdirectories per email
- Subdirectory format: `YYYY-MM-DD_NNNN_Subject/`
- Duplicate filenames are handled with numeric suffixes

## Example Output

```
emails/
├── 2024-01-15_0001_Property_inquiry.eml
├── 2024-01-16_0002_Meeting_confirmation.eml
└── 2024-01-17_0003_Invoice_attached.eml

attachments/
├── 2024-01-15_0001_Property_inquiry/
│   └── floorplan.pdf
└── 2024-01-17_0003_Invoice_attached/
    ├── invoice.pdf
    └── terms.docx
```

## Error Handling

The script includes comprehensive error handling for:
- Authentication failures
- Network errors
- API rate limiting (includes automatic delays)
- Invalid message IDs
- File system errors

## Logging

- Default: INFO level logging to console
- Use `--verbose` flag for DEBUG level logging
- Logs include timestamps, levels, and detailed messages

## Troubleshooting

### Authentication Errors

If you receive authentication errors:
1. Verify your Tenant ID, Client ID, and Client Secret are correct
2. Ensure the Azure AD application has `Mail.Read` permission
3. Verify admin consent has been granted for the permission

### Search Returns No Results

1. Verify the mailbox email address is correct
2. Check the search query syntax
3. Ensure the mailbox has emails matching the criteria

### Rate Limiting

The script includes built-in delays to avoid rate limiting. If you still encounter rate limit errors:
- The script will log the error
- Try again after a few minutes
- Consider reducing the search scope

## API Permissions Required

| Permission | Type | Description |
|------------|------|-------------|
| Mail.Read | Application | Read mail in all mailboxes |

## Security Notes

- Client secrets should be kept secure and never committed to version control
- Consider using environment variables or Azure Key Vault for credentials
- The client credentials flow is suitable for unattended/daemon scenarios
- Ensure proper access controls on output directories

## License

This is a utility script for personal/organizational use.

## Contributing

Feel free to submit issues or pull requests for improvements.
