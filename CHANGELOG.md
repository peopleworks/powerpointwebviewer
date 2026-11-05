# Changelog

All notable changes to Live Website Viewer for PowerPoint.

## [1.4.0] - 2026-11-05

### Security
- Links saved inside a presentation are now checked before loading. Before, a modified `.pptx` file could make the add-in run code (for example a `javascript:` link). Only `http://` and `https://` websites are opened now, whether typed, saved or opened in a window.
- The add-in now asks for the minimum Office permission (`Restricted` instead of `ReadWriteDocument`). It only needs to save its own URL setting.
- Embedded sites can no longer navigate away from the add-in (`allow-top-navigation` removed from the iframe sandbox).
- Messages are built without `innerHTML` when they contain a URL.

### Fixed
- **Window** button: Office's Dialog API only opens pages on the add-in's own domain, so external sites failed and the backup popup was usually blocked. It now opens the new `dialog.html` bridge page, which redirects to the site. If a window still can't open, a clickable link is shown instead.
- Sites that refuse to be embedded showed the browser's "refused to connect" page with no help. Browsers don't report this, so the viewer now briefly shows a **"Site not showing? Open in window"** shortcut after a page loads.
- YouTube start times like `t=1m30s` or `t=1h2m` now start at the right second (they used to start at 1 or 2 seconds).
- Unlisted Vimeo videos (`vimeo.com/123456/abcdef1234`) now play.
- `music.youtube.com` links are converted to the embeddable player.
- YouTube playlist parameters are validated, so extra player parameters can't be slipped in.
- The manifest now points to PNG icons (`icon-32.png`, `icon-64.png`); Office does not accept SVG icons.
- `SupportUrl` now points to the GitHub issues page.

### Added
- The add-in is inserted at a 16:9 size (800×450) instead of the small default.
- The version number is shown on the welcome screen.
- `url-tools.js`: URL check and YouTube/Vimeo conversion in one file, with unit tests (`node --test tests/url-tools.test.js`).
- GitHub Actions workflow: runs the tests, checks the XML files and checks that the version matches in `manifest.xml` and `viewer.html`.
- `LICENSE` file (MIT).
- Accessibility: `lang` attribute, labels for icon-only buttons, status announcements, reduced-motion support.

### Upgrade notes
- Upload **all** of these to the server: `viewer.html`, `url-tools.js`, `dialog.html`, `icon-32.png`, `icon-64.png`, `icon-80.png`, `icon-128.png`, `manifest.xml`.
- The add-in `Id` is unchanged, so existing presentations keep working.
- Users who installed the add-in from `manifest.xml` should load the new `manifest.xml` to get the new icons and permission level. `viewer.html` updates for everyone automatically.

## [1.3.0] - 2026-05-21
- Video pages fixed (iframe `allow` attribute and sandbox permissions).
- YouTube/Vimeo links auto-converted to their embeddable player.
- Reload and Fullscreen buttons.
- Window button shown on Desktop too.

## [1.2.0] - 2026-03-10
- Fixed blank display in PowerPoint Online.
- Office Dialog API fallback.
- Brand icon.
