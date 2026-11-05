# Live Website Viewer for PowerPoint

[![CI](https://github.com/peopleworks/powerpointwebviewer/actions/workflows/ci.yml/badge.svg)](https://github.com/peopleworks/powerpointwebviewer/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/github/license/peopleworks/powerpointwebviewer?color=blue)](LICENSE)
[![Version](https://img.shields.io/badge/version-1.4.0-2c6cb0)](CHANGELOG.md)
[![PowerPoint](https://img.shields.io/badge/PowerPoint-Online_%7C_Windows-B7472A)](#quick-start)
[![Office Add-in](https://img.shields.io/badge/Office_Add--in-Content-0078D4)](https://learn.microsoft.com/office/dev/add-ins/)
[![Dependencies](https://img.shields.io/badge/dependencies-zero-1e8a4c)](#repository-layout)
[![GitHub stars](https://img.shields.io/github/stars/peopleworks/powerpointwebviewer?style=social)](https://github.com/peopleworks/powerpointwebviewer/stargazers)

**🇪🇸 [Leer en español](README.es.md)**

<img alt="Live websites, inside your slides. A PowerPoint slide showing a live sales dashboard through the add-in, with the URL bar, Load and Window buttons on top." src="assets/readme/cover-en.svg" width="100%">

**Put a real, live website on a PowerPoint slide.** Dashboards that keep updating, documentation
you can scroll, a YouTube video that plays in place — during the presentation, without switching
windows and without screenshots that are out of date the moment you take them.

It is a PowerPoint **content add-in**: a box you place on a slide, like a picture or a chart. Each
box remembers its own URL inside the `.pptx`, so the presentation opens exactly as you left it.

**Created by Pedro Hernandez** — [PeopleWorks Services](https://peopleworksservices.com)

---

## Contents

- [Why this exists](#why-this-exists)
- [What you get](#what-you-get)
- [Screenshots](#screenshots)
- [Quick start](#quick-start)
- [How it works](#how-it-works)
- [Presenting with it](#presenting-with-it)
- [YouTube and Vimeo](#youtube-and-vimeo)
- [Why some websites do not appear](#why-some-websites-do-not-appear)
- [Security and privacy](#security-and-privacy)
- [Customize and self-host](#customize-and-self-host)
- [Development](#development)
- [Troubleshooting](#troubleshooting)
- [Status and roadmap](#status-and-roadmap)
- [Repository layout](#repository-layout)

---

## Why this exists

Every presenter who shows something *live* knows the routine: leave the slideshow, find the
browser, find the tab, zoom in, come back, lose your place. The alternative — a screenshot — is
already stale while you paste it.

PowerPoint can embed a web page, but not out of the box. This add-in is the missing piece, kept
deliberately small: a handful of static files, no build, no server code, no account. You can read
the whole thing in an afternoon and host it anywhere that serves HTTPS.

| You want to show… | Without the add-in | With the add-in |
| --- | --- | --- |
| A dashboard (Power BI page, Grafana, status page) | Screenshot, already old | **Live**, updating on the slide |
| A YouTube or Vimeo video | Leave the slideshow, find the tab | **Plays in place**, paste the normal link |
| Documentation, a demo app, an internal tool | Alt-Tab and hope | **Scroll and click** inside the slide |

## What you get

| | Feature | |
| --- | --- | --- |
| 🌐 | **Live websites on slides** | Any site that allows embedding, fully interactive |
| 🎯 | **One URL per slide** | Saved inside the presentation, restored when it opens |
| 🎬 | **Video ready** | YouTube and Vimeo links become their embeddable player automatically |
| 🪟 | **Window fallback** | Sites that refuse embedding open in a separate window, in one click |
| 🎤 | **Presentation mode** | Hide the controls, reload, go full screen |
| 🛡️ | **Safe by default** | Only `http://` and `https://` links, even ones saved in a file someone sent you |
| ⚡ | **Zero dependencies** | Static HTML + one small JS file + an XML manifest |
| 🧪 | **Tested** | Unit tests for the URL rules, run on every push |

## Screenshots

These are real screenshots of `viewer.html`, rendered from this repository by
[`tools/readme-assets/screenshots.js`](tools/readme-assets/screenshots.js). The dashboard is a
fictional demo page.

| A live dashboard on the slide | Presenting: controls hidden |
| --- | --- |
| <img alt="The add-in showing a live Q3 sales dashboard, with the URL bar and buttons on top." src="assets/readme/viewer-loaded.png"> | <img alt="The same dashboard with the controls hidden; only a small gear button remains in the corner." src="assets/readme/viewer-presenting.png"> |
| **"Site not showing?" shortcut after each load** | **A suspicious saved link is refused** |
| <img alt="A short bar at the bottom of the slide offers: Site not showing? Open in window." src="assets/readme/viewer-hint.png"> | <img alt="A warning: Saved link not opened. This slide contains a saved link that is not a normal website address." src="assets/readme/viewer-safety.png"> |
| **Welcome screen** | **If a window is blocked, a link to click** |
| <img alt="The welcome screen of the add-in, asking for a URL, with the version number." src="assets/readme/viewer-welcome.png"> | <img alt="A message explaining the window could not open automatically, with a clickable link to the site." src="assets/readme/viewer-open-link.png"> |

## Quick start

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="assets/readme/install-paths-dark.svg">
  <img alt="Three ways to install: PowerPoint Online by uploading manifest.xml; PowerPoint Desktop through a trusted shared folder; or self-hosting the files on your own HTTPS server." src="assets/readme/install-paths-light.svg">
</picture>

The add-in is already hosted at `https://powerpointwebviewer.peopleworksservices.com/viewer.html`,
so for options A and B you only need [`manifest.xml`](manifest.xml).

### Option A — PowerPoint Online (easiest)

1. Download [`manifest.xml`](manifest.xml).
2. Open a presentation in [PowerPoint for the web](https://www.office.com).
3. **Insert → Add-ins → Upload My Add-in**, and choose `manifest.xml`.
4. Insert the add-in on a slide, type a URL, and press **Enter**.

### Option B — PowerPoint Desktop (Windows)

PowerPoint Desktop has no "Upload" button. It reads add-ins from a **trusted shared folder**.

<details>
<summary><strong>Step-by-step: shared folder catalog</strong></summary>

**1. Create a shared folder for manifests**

1. Create a folder, for example `C:\AddinManifests`, and copy `manifest.xml` into it.
2. Right-click the folder → **Properties** → **Sharing** → **Share…**
3. Add your user (or *Everyone*) with **Read** permission.
4. Note the network path, for example `\\YOUR-PC\AddinManifests`.

**2. Register the folder in PowerPoint**

1. **File → Options → Trust Center → Trust Center Settings… → Trusted Add-in Catalogs**
2. Paste the network path in **Catalog URL** and click **Add catalog**.
3. Tick **Show in Menu**, then **OK → OK**.
4. **Close and reopen PowerPoint** (required).

**3. Insert the add-in**

1. **Home → Add-ins → Advanced** (at the bottom).
2. Choose **SHARED FOLDER** at the top.
3. Select **Live Website Viewer** and click **Add**.

</details>

### Option C — Host it yourself

1. Copy `viewer.html`, `url-tools.js`, `dialog.html` and the `icon-*.png` files to any HTTPS site,
   in the same folder.
2. Edit `manifest.xml` so the URLs point to your site ([details below](#1-point-the-manifest-to-your-server)).
3. Load your `manifest.xml` with option A or B.

> **Updating?** `viewer.html` updates for everyone as soon as the server has the new version.
> Re-load `manifest.xml` only to pick up manifest changes such as new icons or permissions.

## How it works

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="assets/readme/how-it-works-dark.svg">
  <img alt="A slide in the .pptx holds the add-in. Office loads viewer.html from your HTTPS host, which checks the URL and converts YouTube and Vimeo links. The website is shown live in a frame. If a site refuses to be framed, the Window button opens dialog.html on the same domain, which then opens the site in its own window." src="assets/readme/how-it-works-light.svg">
</picture>

1. **The slide holds the add-in.** PowerPoint stores a reference to it in the `.pptx`, plus the
   add-in's own settings — here, a single value: `webViewerUrl`.
2. **Office loads `viewer.html`** from the URL in the manifest, inside the box on the slide, and
   gives it [Office.js](https://learn.microsoft.com/office/dev/add-ins/).
3. **The viewer checks the URL** with `safeUrl()` and converts video links with `toEmbedUrl()`
   (both in [`url-tools.js`](url-tools.js)), then shows the site in a sandboxed `<iframe>`.
4. **The URL is saved** with `Office.context.document.settings`, so it travels with the file.
5. **The Window button** opens `dialog.html` through Office's Dialog API. That page must be on the
   add-in's own domain; it checks the URL again and redirects to the site.

## Presenting with it

| Control | What it does |
| --- | --- |
| **URL box + Enter / Load** | Loads the site. `example.com` becomes `https://example.com/`. |
| **↻ Reload** | Reloads the current site — handy for a dashboard before you present it. |
| **⛶ Fullscreen** | Fills the screen where the host allows it; otherwise hides the controls. |
| **Window** | Opens the current site in a separate window. |
| **✕ Hide controls** | Clean view for the audience. A small **⚙** in the corner brings them back. |

Tips for a smooth talk:

- **Load and check every slide before you present.** The add-in saves the URL, not the page.
- **Hide the controls** on slides where the audience should only see the site.
- **Resize generously.** The add-in is inserted at 16:9 (800×450); drag the blue border to fit.
- In **PowerPoint Online**, the yellow **POWERPOINT ONLINE** badge reminds you that more sites
  refuse embedding there; keep the **Window** button in mind.

## YouTube and Vimeo

Paste the normal link from the address bar or the **Share** button. The viewer turns it into the
embeddable player, which is the only form those sites allow inside a frame.

| You paste | The slide plays |
| --- | --- |
| `youtube.com/watch?v=ID` · `m.youtube.com/…` · `music.youtube.com/…` | `youtube.com/embed/ID` |
| `youtu.be/ID` · `youtube.com/shorts/ID` · `youtube.com/live/ID` | `youtube.com/embed/ID` |
| `…&t=90` · `t=90s` · `t=1m30s` · `t=1h2m3s` | Starts at that second (`?start=90`) |
| `youtube.com/watch?v=ID&list=PL…` | The video, inside its playlist |
| `youtube.com/playlist?list=PL…` | The whole playlist |
| `vimeo.com/123456` · `vimeo.com/channels/…/123456` · `vimeo.com/groups/…/videos/123456` | `player.vimeo.com/video/123456` |
| `vimeo.com/123456/abcdef1234` (unlisted) | `player.vimeo.com/video/123456?h=abcdef1234` |

Lookalike domains such as `youtube.com.example.com` are **not** converted. Every rule is covered by
[`tests/url-tools.test.js`](tests/url-tools.test.js).

## Why some websites do not appear

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="assets/readme/embedding-dark.svg">
  <img alt="Sites that allow framing are shown live on the slide. Sites that send X-Frame-Options or a frame-ancestors policy are blocked by the browser with a refused-to-connect page; the add-in then offers a Site not showing? Open in window shortcut." src="assets/readme/embedding-light.svg">
</picture>

**The website decides, not the add-in.** A site can tell browsers "never show me inside a frame"
with the `X-Frame-Options` header or a `Content-Security-Policy: frame-ancestors` rule. Browsers
obey it everywhere — on any web page, not just in PowerPoint — and no add-in can override it.

What makes this tricky: **the browser does not tell the page that it happened.** The frame simply
shows the browser's own "refused to connect" page. So instead of guessing, the viewer shows a short
**"Site not showing? Open in window"** bar about 2.5 seconds after each page loads, and hides it
again after 15 seconds. It is skipped for YouTube and Vimeo (they always allow their player) and
while the controls are hidden.

| Usually works | Usually refuses |
| --- | --- |
| Wikipedia, documentation sites, Grafana and most dashboards, your own internal tools, YouTube and Vimeo videos | Google search, Gmail, Facebook, X/Twitter, LinkedIn, most banks |

If a site **you own** refuses, add the add-in's domain to its `frame-ancestors` policy (for PowerPoint
for the web, the Office domains too), and remove `X-Frame-Options` if it sends one.

## Security and privacy

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="assets/readme/url-safety-dark.svg">
  <img alt="URLs typed in the toolbar, saved inside the .pptx, or opened with the Window button all pass through safeUrl. Anything that is not http or https is refused with a friendly message. Accepted URLs go through toEmbedUrl and are shown on the slide." src="assets/readme/url-safety-light.svg">
</picture>

**Every URL passes the same check.** That includes the URL saved inside the presentation: a
`.pptx` is a ZIP file anyone can edit, so a file someone sends you is treated like untrusted input.

| Protection | Detail |
| --- | --- |
| **Only `http://` and `https://`** | `javascript:`, `data:`, `file:`, `blob:` and every other scheme are refused. Fixed in 1.4.0. |
| **No embedded credentials** | `https://user:password@host` is refused. |
| **Minimum Office permission** | The manifest asks for `Restricted`: the add-in can store its own setting and nothing else in your document. |
| **Sandboxed frame** | The site runs in a sandboxed `<iframe>` and cannot navigate the add-in away (`allow-top-navigation` is not granted). |
| **No HTML from URLs** | Messages that show a URL are built as text, never as HTML. |
| **No tracking** | The add-in sends nothing anywhere. The only request is the site you ask it to show. |

<details>
<summary><strong>Every permission the frame gets, and why</strong></summary>

`sandbox` flags:

| Flag | Why it is needed |
| --- | --- |
| `allow-scripts` | Almost every modern site and every video player runs JavaScript. |
| `allow-same-origin` | Lets the site keep **its own** origin: cookies, logins, video players. Without it YouTube does not play (it was removed once, in 1.2.x, and broke video). It is safe here because only `http(s)` sites are ever loaded. |
| `allow-forms` | Search boxes and filters on dashboards. |
| `allow-popups`, `allow-popups-to-escape-sandbox` | Links such as "Watch on YouTube" open a normal, un-sandboxed tab. |
| `allow-presentation`, `allow-modals` | Cast buttons, and `alert()`/`confirm()` dialogs some sites use. |
| `allow-downloads` | Exports and downloads from dashboards. |

`allow` (feature policy): `autoplay; encrypted-media; fullscreen; picture-in-picture; clipboard-write; accelerometer; gyroscope; web-share` — what YouTube and Vimeo ask for.

Deliberately **not** set: a restrictive `referrerpolicy` (`no-referrer` makes YouTube refuse to
play) and a restrictive `frame-src` content policy (it would block the very sites you want to show).

</details>

**What is stored in your presentation:** one value, the URL of each add-in box. Do not use links
that contain passwords, tokens or private identifiers — anyone who receives the file can read them.

Found a vulnerability? See [SECURITY.md](SECURITY.md).

## Customize and self-host

Want your own name, colors or domain? Everything is in plain files.

### 1. Point the manifest to your server

In `manifest.xml`:

```xml
<IconUrl DefaultValue="https://YOUR-DOMAIN.com/icon-32.png"/>
<HighResolutionIconUrl DefaultValue="https://YOUR-DOMAIN.com/icon-64.png"/>
<SupportUrl DefaultValue="https://YOUR-DOMAIN.com/support"/>

<SourceLocation DefaultValue="https://YOUR-DOMAIN.com/viewer.html"/>
```

`viewer.html`, `url-tools.js` and `dialog.html` **must be in the same folder** on the same domain:
the Window button only works when `dialog.html` shares the add-in's domain.

### 2. Give it your own identity

```xml
<Id>YOUR-UNIQUE-GUID-HERE</Id>
<ProviderName>Your Name or Company</ProviderName>

<DisplayName DefaultValue="Your Custom Name"/>
<Description DefaultValue="Your custom description."/>
```

> Generate a GUID once (for example with `[guid]::NewGuid()` in PowerShell) and **never change it
> afterwards**: presentations that already use your add-in find it by this Id.

### 3. Change the branding

In `viewer.html`, the footer near the bottom of `<body>`:

```html
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

And the welcome screen, built by `showWelcome()` in the `<script>` section:

```js
var credit = make('div', 'credit');
credit.innerHTML = 'Created by <strong>Your Name</strong> &mdash; Your Company &middot; v' + APP_VERSION;
```

To change the icon, edit `icon.svg` and export it as `icon-32.png`, `icon-64.png`, `icon-80.png`
and `icon-128.png` (Office does not accept SVG icons).

### 4. Change the colors

| Color | Where in `viewer.html` | What it does |
| --- | --- | --- |
| `#4a90d9` / `#2c6cb0` | `body` gradient, `.btn-primary`, `.spinner` | Blue border and primary accents |
| `#27ae60` | `.btn-window`, `.overlay-btn`, `#hint .hint-open` | Window / open buttons |
| `#1a1a2e` / `#16213e` | `#branding` gradient | Footer background |
| `#63b3ed` | `.brand-name`, links in the footer | Brand highlight |
| `#f0f2f5` | `#appWrapper`, `#overlay` | Background |

### 5. Host it

Any static HTTPS host works. Office add-ins **require HTTPS** with a valid certificate.

<details>
<summary><strong>IIS (Windows Server)</strong></summary>

- Copy `viewer.html`, `url-tools.js`, `dialog.html`, the `icon-*.png` files, `manifest.xml` and
  [`web.config`](web.config) to the site folder.
- `web.config` redirects HTTP to HTTPS and sets MIME types and basic headers.
- Install the [URL Rewrite](https://www.iis.net/downloads/microsoft/url-rewrite) module if needed.

</details>

<details>
<summary><strong>Apache</strong></summary>

```apache
# .htaccess
RewriteEngine On
RewriteCond %{HTTPS} off
RewriteRule ^(.*)$ https://%{HTTP_HOST}/$1 [R=301,L]
```

</details>

<details>
<summary><strong>Nginx</strong></summary>

```nginx
server {
    listen 80;
    server_name your-domain.com;
    return 301 https://$host$request_uri;
}

server {
    listen 443 ssl;
    server_name your-domain.com;

    location / {
        root /var/www/powerpointwebviewer;
        index viewer.html;
    }
}
```

</details>

<details>
<summary><strong>GitHub Pages (free)</strong></summary>

1. Fork this repository.
2. **Settings → Pages**, publish from the `main` branch.
3. Your viewer is at `https://YOUR-USERNAME.github.io/powerpointwebviewer/viewer.html`.
4. Update `manifest.xml` with that URL and load it.

</details>

> **Do not** add `X-Frame-Options` or a `frame-ancestors` policy to the server that hosts the
> add-in: PowerPoint itself has to embed `viewer.html`.

## Development

No build step and no packages to install. You need [Node.js](https://nodejs.org/) 18+ only to run
the tests.

```bash
node --test tests/url-tools.test.js
```

**Try the viewer in a browser.** Serve the folder over HTTP and open `viewer.html`. Outside Office
it runs in *standalone* mode: everything works except saving the URL and the Office dialog.

```bash
npx --yes http-server -p 8080
```

**Try it inside PowerPoint.** Office needs HTTPS. Host your branch somewhere with a certificate
(GitHub Pages works), point a copy of `manifest.xml` at it, and sideload that copy.

**Continuous integration.** [`.github/workflows/ci.yml`](.github/workflows/ci.yml) runs on every
push and pull request:

1. the unit tests,
2. a check that `manifest.xml` and `web.config` are well-formed XML,
3. a check that the version in `manifest.xml` matches `APP_VERSION` in `viewer.html`.

**README images.** The diagrams are generated, so they stay consistent in light/dark and
English/Spanish:

```bash
node tools/readme-assets/build-diagrams.js    # SVG diagrams and covers
node tools/readme-assets/screenshots.js       # PNG screenshots (needs Playwright)
```

Contributions are welcome — see [CONTRIBUTING.md](CONTRIBUTING.md).

## Troubleshooting

<details>
<summary><strong>The website shows blank or "refused to connect"</strong></summary>

The site does not allow embedding (see [why](#why-some-websites-do-not-appear)). Click
**Open in window** on the bar that appears after loading, or the **Window** button.

</details>

<details>
<summary><strong>The Window button shows a link instead of opening a window</strong></summary>

The browser or Office blocked the window. Click the link: it opens the site in a new tab.
If you self-host, check that `dialog.html` is in the same folder and domain as `viewer.html`.

</details>

<details>
<summary><strong>"Saved link not opened"</strong></summary>

The URL stored in that slide is not a normal `http(s)` address, so it was not loaded. This
happens with presentations edited outside PowerPoint. Type the correct URL and press **Enter**;
it replaces the stored one.

</details>

<details>
<summary><strong>A video does not play</strong></summary>

- Use the normal YouTube or Vimeo link; the add-in converts it.
- The owner may have disabled embedding for that video. Try it on YouTube with **Share → Embed**.
- Make sure the server has version 1.3.0 or later of `viewer.html` (the version is shown on the
  welcome screen).

</details>

<details>
<summary><strong>The add-in disappears from the ribbon (Desktop)</strong></summary>

Add-ins loaded from a shared folder are not pinned to the ribbon. Insert it again from
**Home → Add-ins → Advanced → SHARED FOLDER**. Slides that already have it keep it: the add-in
is saved with the presentation.

</details>

<details>
<summary><strong>I can't select or resize the add-in after making it bigger (Online)</strong></summary>

A known PowerPoint for the web limitation with content add-ins. Click outside the add-in, then
click its **blue border** (not the website). If that fails, **Ctrl+Z** or re-insert it.

</details>

<details>
<summary><strong>Nothing loads at all</strong></summary>

Check that the add-in's server answers over HTTPS, and that `url-tools.js` is next to
`viewer.html` — without it the viewer cannot start.

</details>

Still stuck? [Open an issue](https://github.com/peopleworks/powerpointwebviewer/issues).

## Status and roadmap

**v1.4.0** — see the [changelog](CHANGELOG.md).

| | |
| --- | --- |
| ✅ | Live websites on slides, one URL per slide, saved in the `.pptx` |
| ✅ | YouTube and Vimeo auto-embed, with start times, playlists and unlisted videos |
| ✅ | Reload, fullscreen, hide controls for presenting |
| ✅ | Window button through Office's Dialog API, with a clickable-link fallback |
| ✅ | "Site not showing?" shortcut for sites that refuse embedding |
| ✅ | URL safety checks, minimum Office permission, sandboxed frame |
| ✅ | Unit tests and CI |
| 🔜 | Zoom / fit-to-slide, so desktop-sized dashboards fit the box |
| 🔜 | Auto-refresh every N minutes, for dashboards |
| 🔜 | Interface in Spanish, chosen from Office's display language |
| 🔜 | Microsoft Marketplace listing, for one-click install |

## Repository layout

```
viewer.html                The add-in page PowerPoint loads inside the slide
url-tools.js               safeUrl() and toEmbedUrl(), shared by viewer.html and dialog.html
dialog.html                Bridge page for the Window button (must share the add-in's domain)
manifest.xml               Office add-in manifest
icon.svg                   Icon source; icon-32/64/80/128.png are exported from it
web.config                 IIS hosting configuration
powerpointwebviewer.html   Hosted documentation page (English / Spanish)
tests/                     Unit tests (node --test)
tools/readme-assets/       Generators for the README diagrams and screenshots
assets/readme/             README images
.github/workflows/ci.yml   Continuous integration
```

## License

[MIT](LICENSE) — free to use, modify and distribute.

## Author

**Pedro Hernandez** — [PeopleWorks Services](https://peopleworksservices.com)

If this add-in saves you an Alt-Tab during a talk, a ⭐ on the repository helps others find it.
