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

### Option A: PowerPoint Online (easiest)

1. Download [`manifest.xml`](manifest.xml)
2. Go to [PowerPoint Online](https://www.office.com) and open a presentation
3. Go to **Insert > Add-ins > Upload My Add-in**
4. Select the `manifest.xml` file
5. Insert the add-in into any slide
6. Enter a URL and press **Enter**

The viewer is hosted at `https://powerpointwebviewer.peopleworksservices.com/viewer.html` — no server setup needed.

### Option B: PowerPoint Desktop (Windows)

> **Important:** PowerPoint Desktop does **not** have an "Upload" button. You need to set up a shared folder catalog.

**Step 1: Create a shared folder for manifests**

1. Create a folder on your PC, e.g. `C:\AddinManifests`
2. Copy `manifest.xml` into that folder
3. Right-click the folder > **Properties** > **Sharing** tab > **Share...**
4. Add your user (or "Everyone") with **Read** permission
5. Note the network path (e.g. `\\YOUR-PC\AddinManifests`)

**Step 2: Register the folder in PowerPoint**

1. Open PowerPoint > **File** > **Options**
2. Click **Trust Center** > **Trust Center Settings...**
3. Click **Trusted Add-in Catalogs**
4. In **Catalog URL**, paste the network path: `\\YOUR-PC\AddinManifests`
5. Click **Add catalog**
6. Check the **Show in Menu** checkbox
7. Click **OK** > **OK**
8. **Close and reopen PowerPoint** (required)

**Step 3: Load the add-in**

1. Go to **Home** > **Add-ins** (in the ribbon)
2. Click **Advanced** (at the bottom)
3. Select **SHARED FOLDER** at the top
4. Select "Live Website Viewer" and click **Add**

### Option C: Self-host

1. Clone this repository
2. Host `viewer.html` on your own HTTPS server
3. Edit `manifest.xml` and update the `SourceLocation` URL to point to your server
4. Follow Option A or B to load the modified manifest

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
