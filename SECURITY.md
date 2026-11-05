# Security policy

## Supported versions

| Version | Supported |
| --- | --- |
| 1.4.x | ✅ |
| 1.3.x and older | ❌ — please update: 1.4.0 fixes a security issue with saved links |

The hosted add-in at `powerpointwebviewer.peopleworksservices.com` always runs the latest version.
If you host your own copy, update `viewer.html`, `url-tools.js` and `dialog.html` together.

## Reporting a vulnerability

**Please do not open a public issue.** Report it privately through GitHub:
**Security → Report a vulnerability** on this repository
([direct link](https://github.com/peopleworks/powerpointwebviewer/security/advisories/new)).

Include, if you can:

- what an attacker could do, and what they need (for example "a crafted `.pptx` file");
- steps to reproduce, or a sample file;
- the PowerPoint version and platform (Online, Windows, Mac).

You will get an answer as soon as possible. Once a fix is published, you will be credited in the
changelog unless you prefer otherwise.

*También puedes reportarlo en español.*

## What is in scope

- Code running in the add-in: `viewer.html`, `url-tools.js`, `dialog.html`.
- The manifest and the hosting configuration in this repository.
- Anything that lets a presentation file, a URL or an embedded site run code with the add-in's
  permissions, read other documents, or navigate the add-in to another page.

## What is not

- **Sites refusing to be embedded** (`X-Frame-Options`, `frame-ancestors`). That is the site's
  choice and the browser enforcing it; see the README.
- **What a website does once you choose to show it.** The add-in shows the site you ask for, inside
  a sandboxed frame; the site's own behavior is the site's responsibility.
- Issues in PowerPoint or Office.js themselves — please report those to Microsoft.

## How the add-in protects you

See [Security and privacy](README.md#security-and-privacy) in the README for the full list: URL
checks on every entry point, the minimum Office permission, the sandboxed frame and what is stored
in your presentation.
