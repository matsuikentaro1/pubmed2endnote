# PubMed2EndNote

Chrome extension that copies a PubMed article to the clipboard as an **EndNote-compatible Word citation field** in one click.

**Install from Chrome Web Store:**  
https://chromewebstore.google.com/detail/pubmed2endnote/fjcjpmjhmlkigknkfcchbbbacgchgofi

## What it does

- Adds a button on PubMed article pages
- Copies an EndNote Cite While You Write (CWYW) citation field to the clipboard
- Paste into Microsoft Word (`Ctrl+V` / `Cmd+V`); EndNote can treat it as a real citation
- Pulls metadata via the official NCBI E-utilities API (authors, journal, volume/issue/pages, DOI, abstract, MeSH, etc.)
- Supports special characters (e.g. umlauts)
- Optional yellow highlight on paste so you can spot new citations quickly

## Requirements

- Google Chrome (or another Chromium browser that can load this extension)
- Microsoft Word + EndNote (Cite While You Write)
- Internet access (for NCBI API)

## Install

### Option A — Chrome Web Store (recommended)

1. Open the [Chrome Web Store listing](https://chromewebstore.google.com/detail/pubmed2endnote/fjcjpmjhmlkigknkfcchbbbacgchgofi)
2. Click **Add to Chrome**

### Option B — Load from this repository

1. Clone or download this repo
2. Open `chrome://extensions/`
3. Enable **Developer mode**
4. **Load unpacked** → select the `PubMed2EndNote_3.0.0` folder

## Usage

1. Open an article on [PubMed](https://pubmed.ncbi.nlm.nih.gov/)
2. Click the blue button (top-right of the page)
3. On first use, set an email address in the options page (see below)
4. Paste into Word
5. In the EndNote tab, run **Update Citations and Bibliography**

### Options

Right-click the blue button → **Settings**, or open extension options from `chrome://extensions/`.

- **Font / size** for the pasted citation (default: Times New Roman)
- **Email** for NCBI E-utilities contact policy (stored only in your browser; sent only with NCBI API requests; not collected by the developer)

See [PRIVACY.md](PRIVACY.md) for details.

## How it works

1. Fetch article XML from NCBI E-utilities
2. Build an EndNote Traveling Library field (`ADDIN EN.CITE`)
3. Write Word-compatible HTML to the clipboard
4. Word rebuilds the field on paste; EndNote recognizes the citation

Everything runs in the browser. No native host or extra desktop installer.

## Troubleshooting

| Problem | What to try |
|--------|-------------|
| No blue button | Reload the page; confirm the extension is enabled |
| Click opens settings | Save an email in options, then try again |
| EndNote does not pick it up | EndNote tab → **Update Citations and Bibliography** |
| Paste is plain text only | Use paste option **Keep Source Formatting** |
| No yellow highlight | Word → File → Options → Advanced → Cut, copy, and paste → set **Pasting from other programs** to **Keep Source Formatting** (Merge Formatting can strip highlight only) |

## Uninstall

Remove the extension from `chrome://extensions/`.

## Issues

Bug reports and ideas: [GitHub Issues](https://github.com/matsuikentaro1/pubmed2endnote/issues)

## License / notice

Free to use. Bibliographic data comes from NCBI. This project is not affiliated with NCBI, NLM, Clarivate, or EndNote.
