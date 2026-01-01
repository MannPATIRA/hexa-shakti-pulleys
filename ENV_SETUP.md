# Environment Setup

## Creating .env File

Create a `.env` file in the root directory with the following content:

```
# Google Sheets Configuration
SPREADSHEET_ID=1nXUr6oQiH2ji5zKCuI1e35JchtKtjJdUMaeforHmyns
SERVICE_ACCOUNT_FILE=sheets-api-473619-dc7d5f869aeb.json
```

## Quick Setup

Run this command in the terminal:

```bash
cat > .env << 'EOF'
# Google Sheets Configuration
SPREADSHEET_ID=1nXUr6oQiH2ji5zKCuI1e35JchtKtjJdUMaeforHmyns
SERVICE_ACCOUNT_FILE=sheets-api-473619-dc7d5f869aeb.json
EOF
```

Or manually create the file and paste the content above.

## Security Note

The `.env` file is already added to `.gitignore` to prevent committing secrets to version control.

