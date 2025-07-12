# PubMed2EndNote Browser Extension

A Chrome/Edge browser extension that exports PubMed citations to clipboard in EndNote-compatible RTF format for Microsoft Word.

## Features

- **One-click export**: Floating icon appears on PubMed article pages for instant citation export
- **Manual PMID entry**: Enter any PMID directly through the extension popup
- **Email integration**: Uses your NCBI account email for API compliance
- **EndNote compatible**: Generates RTF format that works directly with Word's EndNote plugin

## Installation

### Developer Mode Installation

1. **Download and extract files**
   ```
   PubMed2EndNote/
   ├── manifest.json
   ├── content.js
   ├── background.js
   ├── options.html
   ├── options.js
   ├── icons/
   │   ├── icon16.png
   │   ├── icon48.png
   │   └── icon128.png
   └── README.md
   ```

2. **Create icon files**
   - Create `icons/` folder
   - Add 16x16, 48x48, and 128x128 pixel PNG icons
   - Temporary: rename any image files to these names for testing

3. **Load extension in Chrome/Edge**
   - Chrome: Navigate to `chrome://extensions/`
   - Edge: Navigate to `edge://extensions/`
   - Enable "Developer mode"
   - Click "Load unpacked"
   - Select the `PubMed2EndNote` folder

## Setup & Usage

### Simple User Flow

1. **Navigate to any PubMed article page** (e.g., `https://pubmed.ncbi.nlm.nih.gov/34886562/`)
2. **Click the floating blue icon** in the top-right corner
3. **First time only**: Settings page opens for NCBI email configuration
4. **Every time after**: Citation is instantly copied to clipboard

### Email Configuration (First Time Only)

When you first click the floating icon:
1. A settings page opens automatically
2. **Find your NCBI email**:
   - Sign in to PubMed with your NCBI account
   - Visit [NCBI Account Settings](https://account.ncbi.nlm.nih.gov/settings/)
   - Copy the email address from your account profile
3. Enter the email and click "Save Settings"

**Why needed?** NCBI requires user identification for API access compliance.

### Daily Usage

- **Just click the floating icon** on any PubMed article page
- Citation is automatically copied to clipboard in EndNote format

### Settings Access (Rare)

If you need to change your email address later:
1. Go to `chrome://extensions/`
2. Find "PubMed2EndNote" 
3. Click "Extension options" account
   - Visit [NCBI Account Settings](https://account.ncbi.nlm.nih.gov/settings/)
   - Copy the email address shown in your account profile
4. Paste the email address in the extension popup
5. Click "Save"

**Why is this required?** NCBI requires user identification for API access to comply with their usage policies and rate limiting.

### 2. Using the Extension

#### Method 1: Floating Icon (Recommended)
1. Navigate to any PubMed article page (e.g., `https://pubmed.ncbi.nlm.nih.gov/34886562/`)
2. Look for the floating blue icon in the top-right corner
3. Click the icon to export the citation
4. If email isn't configured, you'll be prompted to set it up

#### Method 2: Manual PMID Entry
1. Click the extension icon in your browser toolbar
2. Enter a PMID in the "Manual PMID Entry" field
3. Click "Export"

### Using in Microsoft Word

1. Place cursor where you want the citation
2. Press `Ctrl+V` (Windows) or `Cmd+V` (Mac) to paste
3. Go to EndNote tab → "Update Citations"
4. Citation formats automatically according to your style

## User Flow Summary

```
🔵 Click floating icon → ⚙️ Setup email (first time) → 📋 Auto-copy → 📝 Paste in Word
```

**That's it!** One click on PubMed articles, paste in Word.

## Supported Pages

- **Individual article pages only**: `https://pubmed.ncbi.nlm.nih.gov/[PMID]/`
- Search result pages are not supported (navigate to individual articles)

## Technical Specifications

### API Integration
- **PubMed eUtils API**: `https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi`
- **Data format**: MEDLINE/NBIB format retrieval
- **Output format**: EndNote XML → RTF format conversion

### Data Flow
1. **PMID extraction**: Extract from PubMed article page URL
2. **API request**: Fetch NBIB data using user's email for NCBI compliance
3. **Format conversion**: Convert NBIB → EndNote XML → RTF format
4. **Clipboard**: Save RTF format to clipboard for Word compatibility

### File Structure

| File | Purpose |
|------|---------|
| `manifest.json` | Extension configuration and permissions |
| `content.js` | Displays floating icon on PubMed article pages |
| `background.js` | Handles API calls and data conversion |
| `options.html/js` | Settings page for email configuration |
## Troubleshooting

### Common Issues

1. **Floating icon doesn't appear**
   - Ensure you're on a PubMed individual article page (`/[number]/` format)
   - Refresh the page
   - Check that the extension is enabled in `chrome://extensions/`

2. **Email configuration issues**
   - Make sure you're signed in to your NCBI account
   - Visit [NCBI Account Settings](https://account.ncbi.nlm.nih.gov/settings/) to verify your email
   - The email must match your NCBI account exactly

3. **Export fails**
   - Verify your email is configured correctly in extension options
   - Check your internet connection
   - Try again after a few seconds (API rate limiting)

4. **Clipboard copy fails**
   - Check browser permissions for clipboard access
   - Ensure you're on an HTTPS page
   - Check for conflicts with other clipboard extensions

5. **EndNote doesn't recognize the citation**
   - Confirm Word's EndNote plugin is enabled
   - Try pasting as RTF format
   - Run "Update Citations" from the EndNote tab

### Error Messages

- **"Please configure your email address first"**: Set up your NCBI email in the popup
- **"PubMed API error"**: API connection issue, wait and retry
- **"No data found for PMID"**: Invalid PMID or article not found
- **"Clipboard copy failed"**: Browser permissions or security settings issue

## Development & Customization

### Development Setup

1. Clone or download the repository
2. Make code modifications as needed
3. Go to `chrome://extensions/` and click "Reload" on the extension

### Customizable Options

- **Citation format**: Modify the `convertToEndNoteRTF` function in `background.js`
- **Icon design**: Replace files in `icons/` folder
- **Icon position**: Adjust `position` styles in `content.js`
- **UI styling**: Modify `popup.css` for popup appearance

## Privacy & Security

- **Email storage**: Your email is stored locally in Chrome's sync storage
- **API usage**: Only used for PubMed API requests with proper identification
- **No tracking**: Extension doesn't collect or transmit personal data
- **Permissions**: Only requests necessary permissions for clipboard and storage access

## License

This project is provided under the MIT License.

## Version History

- **v1.0**: Initial release
  - Floating icon on PubMed article pages
  - Email configuration requirement
  - Manual PMID entry via popup
  - EndNote RTF format compatibility

## Support

For issues or feature requests, please report them on the GitHub Issues page.