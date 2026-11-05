# Contributing

Thanks for helping! Issues and pull requests are welcome, in English or Spanish.
*¡Gracias por ayudar! Puedes escribir en inglés o en español.*

## The one rule: keep it simple

Most people who use this add-in found it in a video and copied a few files to a server. That
simplicity is the point. Please keep:

- **No build step and no runtime dependencies.** Plain HTML, CSS and JavaScript that runs as-is.
- **Few files.** A new file is fine when it earns its place (like `url-tools.js`, which is shared
  and tested), but say why in the pull request.
- **Readable code.** Someone following a tutorial should be able to read `viewer.html` top to bottom.

## Getting started

```bash
git clone https://github.com/peopleworks/powerpointwebviewer.git
cd powerpointwebviewer
node --test tests/url-tools.test.js
```

To try the viewer in a browser, serve the folder and open `viewer.html`:

```bash
npx --yes http-server -p 8080
```

To try it inside PowerPoint you need HTTPS: host your branch (GitHub Pages works), point a copy of
`manifest.xml` at it, and sideload that copy. See [Development](README.md#development).

## Before you open a pull request

- [ ] `node --test tests/url-tools.test.js` passes.
- [ ] New URL or video rules have a test in `tests/url-tools.test.js`.
- [ ] Every URL that reaches the frame, a window or `dialog.html` goes through `UrlTools.safeUrl()`.
- [ ] Nothing builds HTML from a URL or other user input (use `textContent`, not `innerHTML`).
- [ ] You tried it in PowerPoint for the web and, if you can, PowerPoint Desktop. Say which in the PR.
- [ ] If you changed what users see, update `README.md` **and** `README.es.md`.
- [ ] If it is a release, bump the version in **both** `manifest.xml` (`<Version>`) and
      `viewer.html` (`APP_VERSION`) — CI checks they match — and add a `CHANGELOG.md` entry.
- [ ] If you changed a diagram or the UI, regenerate the images:
      `node tools/readme-assets/build-diagrams.js` and `node tools/readme-assets/screenshots.js`.

## Things to never change

- **The add-in `<Id>` in `manifest.xml`.** Existing presentations find the add-in by it.
- **The settings key `webViewerUrl`.** Saved presentations use it.

## Reporting a security problem

Please do not open a public issue. See [SECURITY.md](SECURITY.md).

## License

By contributing, you agree that your contributions are licensed under the [MIT License](LICENSE).
