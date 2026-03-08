# Live Website Viewer for PowerPoint

Embed live, interactive websites directly inside your PowerPoint presentations. Each slide can display a different website URL.

## Features

- **Live websites inside slides** - Display any website directly in PowerPoint
- **Per-slide URLs** - Each slide remembers its own URL
- **Presentation mode** - Hide/show controls with one click
- **Smart URL handling** - Auto-adds `https://` if you forget the protocol
- **Error handling** - Friendly messages when a site can't be embedded
- **Zero dependencies** - Just HTML, JS, and an XML manifest

## Quick Start

### Option 1: Use the hosted version (easiest)

1. Download [`manifest.xml`](manifest.xml)
2. Open PowerPoint
3. Go to **Insert > My Add-ins > Upload My Add-in**
4. Select the `manifest.xml` file
5. Insert the add-in into any slide
6. Enter a URL and press **Enter**

The viewer is hosted at `https://powerpointwebviewer.peopleworksservices.com/viewer.html` — no server setup needed.

### Option 2: Self-host

1. Clone this repository
2. Host `viewer.html` on your own HTTPS server
3. Edit `manifest.xml` and update the `SourceLocation` URL to point to your server
4. Upload the modified `manifest.xml` into PowerPoint

## Files

| File | Description |
|------|-------------|
| `viewer.html` | The add-in that runs inside PowerPoint slides |
| `manifest.xml` | Office Add-in manifest (registers the add-in with PowerPoint) |
| `powerpointwebviewer.html` | Documentation page (English/Spanish) |
| `web.config` | IIS configuration for Windows Server hosting |

## How It Works

The add-in uses Microsoft's [Office.js](https://learn.microsoft.com/en-us/office/dev/add-ins/) API to:

1. Load a lightweight HTML page inside a PowerPoint content add-in
2. Accept a URL from the user
3. Display the website in an iframe
4. Persist the URL using `Office.context.document.settings` so it's saved with the presentation

## Requirements

- PowerPoint (Desktop or Online)
- HTTPS hosting for the viewer page (already provided at peopleworksservices.com)

## Limitations

- Some websites block iframe embedding (e.g., Google, Facebook) via `X-Frame-Options` or CSP headers
- The viewer requires an internet connection to load websites

## License

MIT License - Free to use, modify, and distribute.

## Author

**PeopleWorks Services**
- Website: [peopleworksservices.com](https://peopleworksservices.com)
