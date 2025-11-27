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

## Azure AD Application Setup

Before you can use this tool, you need to register an application in Azure AD and configure the necessary permissions. Follow these steps:

### Step 1: Register a New Application

1. Sign in to the [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** (or **Microsoft Entra ID**)
3. Click on **App registrations** in the left sidebar
4. Click **+ New registration** at the top
5. Fill in the application details:
   - **Name**: Enter a name (e.g., "Outlook Email Downloader")
   - **Supported account types**: Select "Accounts in this organizational directory only"
   - **Redirect URI**: Leave blank (not needed for client credentials flow)
6. Click **Register**

### Step 2: Note Your Tenant ID and Client ID

After registration, you'll be taken to the app's Overview page:

1. **Copy the Application (client) ID**: This is your `CLIENT_ID`
   - Format: `12345678-1234-1234-1234-123456789012`
   - Save this value for later

2. **Copy the Directory (tenant) ID**: This is your `TENANT_ID`
   - Format: `12345678-1234-1234-1234-123456789012`
   - Save this value for later

### Step 3: Create a Client Secret

1. In your app registration, click on **Certificates & secrets** in the left sidebar
2. Under **Client secrets**, click **+ New client secret**
3. Add a description (e.g., "Email downloader secret")
4. Select an expiration period:
   - 6 months, 12 months, 24 months, or custom
   - **Note**: You'll need to create a new secret when this expires
5. Click **Add**
6. **IMPORTANT**: Copy the **Value** immediately (not the Secret ID!)
   - This is your `CLIENT_SECRET`
   - Format: `AbC1~dEf2gHi3jKl4MnO5pQr6StU7vWx8YzA9BcD0`
   - **This value is only shown once** - if you lose it, you'll need to create a new secret
   - Save this value securely

### Step 4: Add API Permissions

1. In your app registration, click on **API permissions** in the left sidebar
2. Click **+ Add a permission**
3. Select **Microsoft Graph**
4. Select **Application permissions** (not Delegated permissions)
5. Search for "Mail.Read" and expand **Mail**
6. Check the box next to **Mail.Read**
7. Click **Add permissions** at the bottom

### Step 5: Grant Admin Consent

**Important**: Application permissions require admin consent.

1. Still on the **API permissions** page, click **Grant admin consent for [Your Organization]**
2. Click **Yes** to confirm
3. You should see a green checkmark in the "Status" column next to Mail.Read

### Step 6: Verify Configuration

Your API permissions should show:

| API / Permission name | Type | Description | Admin consent | Status |
|-----------------------|------|-------------|---------------|--------|
| Microsoft Graph: Mail.Read | Application | Read mail in all mailboxes | Required | ✓ Granted for [Org] |

### Step 7: Save Your Credentials

Create a `.env` file in your project directory:

```bash
cp .env.example .env
```

Edit `.env` and add your credentials:

```
TENANT_ID=your-tenant-id-from-step-2
CLIENT_ID=your-client-id-from-step-2
CLIENT_SECRET=your-client-secret-from-step-3
MAILBOX_EMAIL=user@yourdomain.com
SEARCH_QUERY=your search terms
```

**Security Note**: Never commit the `.env` file to version control. It's already in `.gitignore`.

## Setup

1. **Create virtual environment** (if not already done):
   ```bash
   mkvirtualenv outlook-downloader -p python3.11
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Verify your credentials**:
   ```bash
   python verify-credentials.py
   ```
   This will test your Azure AD credentials and API access.

## Usage

### Recommended: Use the Comprehensive Downloader

The comprehensive downloader searches all possible locations to maximize email discovery:

```bash
python download-comprehensive.py --verbose
```

This script:
- Searches the main mailbox messages
- Searches archive mailbox (if available)
- Searches special folders (recoverable items, etc.)
- Deduplicates results automatically
- Achieved 95.6% coverage in testing (303 out of 317 emails)

### Alternative: Direct Command-line Usage

If you prefer to specify all parameters on the command line:

```bash
python outlook-downloader.py \
  --account user@example.com \
  --search "searchterm1 OR searchterm2" \
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
--search "example.com"

# Search with OR operator
--search "searchterm1 OR searchterm2"

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
