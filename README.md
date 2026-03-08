# Live Website Viewer for PowerPoint

Embed live, interactive websites directly inside your PowerPoint presentations. Each slide can display a different website URL.

**Created by Pedro Hernandez** — [PeopleWorks Services](https://peopleworksservices.com)

---

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

## Customization Guide

Want to host your own version or personalize it? Here's everything you need to change:

### 1. Change the hosting URL

Edit `manifest.xml` and update these two lines with your server address:

```xml
<IconUrl DefaultValue="https://YOUR-DOMAIN.com/icon.png"/>
<SupportUrl DefaultValue="https://YOUR-DOMAIN.com/support"/>

<SourceLocation DefaultValue="https://YOUR-DOMAIN.com/viewer.html"/>
```

### 2. Change the add-in identity

In `manifest.xml`, generate a new unique ID and update the provider name:

```xml
<Id>YOUR-UNIQUE-GUID-HERE</Id>
<ProviderName>Your Name or Company</ProviderName>

<DisplayName DefaultValue="Your Custom Name"/>
<Description DefaultValue="Your custom description."/>
```

> You can generate a GUID at [guidgenerator.com](https://www.guidgenerator.com/)

### 3. Customize the branding

In `viewer.html`, find the **branding footer** near the bottom of the `<body>` and update:

```html
<!-- Branding footer -->
<div id="branding">
  <div class="brand-left">
    <div class="brand-logo">P</div>                          <!-- Your logo letter -->
    <span class="brand-name">PeopleWorks</span>              <!-- Your brand name -->
    <span class="brand-separator">|</span>
    <span class="brand-author">by Pedro Hernandez</span>     <!-- Your name -->
  </div>
  <div class="brand-right">
    <a href="https://YOUR-DOMAIN.com">YOUR-DOMAIN.com</a>    <!-- Your URL -->
  </div>
</div>
```

Also update the **welcome screen** in the same file:

```html
<!-- Welcome overlay -->
<div id="overlay">
  ...
  <div style="margin-top: 24px; opacity: 0.5; font-size: 12px; color: #999;">
    Created by <strong>Your Name</strong> &mdash; Your Company
  </div>
</div>
```

And the `showWelcome()` function in the `<script>` section — it has the same text.

### 4. Change the colors

In `viewer.html`, the main colors are defined in the CSS:

| Color | Where | What it does |
|-------|-------|-------------|
| `#4a90d9` | `.btn-primary`, `.spinner`, `#urlInput:focus` | Primary blue (buttons, accents) |
| `#1a1a2e` / `#16213e` | `#branding` gradient | Footer background |
| `#63b3ed` | `.brand-name`, `a` in branding | Brand text highlight |
| `#f0f2f5` | `body`, `#overlay` | Background color |

### 5. Deploy on your own server

**IIS (Windows Server):**
- Copy `viewer.html`, `manifest.xml`, and `web.config` to your site folder
- The included `web.config` handles CORS, HTTPS redirect, and MIME types
- Make sure you have a valid SSL certificate (Office Add-ins require HTTPS)
- Install the [URL Rewrite](https://www.iis.net/downloads/microsoft/url-rewrite) module if not already installed

**Apache:**
```apache
# .htaccess
Header set Access-Control-Allow-Origin "*"
RewriteEngine On
RewriteCond %{HTTPS} off
RewriteRule ^(.*)$ https://%{HTTP_HOST}/$1 [R=301,L]
```

**Nginx:**
```nginx
server {
    listen 80;
    server_name your-domain.com;
    return 301 https://$host$request_uri;
}

server {
    listen 443 ssl;
    server_name your-domain.com;

    add_header Access-Control-Allow-Origin *;

    location / {
        root /var/www/powerpointwebviewer;
        index viewer.html;
    }
}
```

**GitHub Pages (free):**
1. Fork this repository
2. Go to **Settings > Pages** and enable GitHub Pages from the `main` branch
3. Your viewer will be at `https://YOUR-USERNAME.github.io/powerpointwebviewer/viewer.html`
4. Update `manifest.xml` with that URL

## Requirements

- PowerPoint (Desktop or Online)
- HTTPS hosting for the viewer page (already provided at peopleworksservices.com)

## Limitations

- Some websites block iframe embedding (e.g., Google, Facebook) via `X-Frame-Options` or CSP headers
- The viewer requires an internet connection to load websites

## Contributing

Contributions are welcome! Feel free to open issues or submit pull requests.

## License

MIT License - Free to use, modify, and distribute.

## Author

**Pedro Hernandez** — [PeopleWorks Services](https://peopleworksservices.com)
