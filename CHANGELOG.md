# Changelog

All notable changes to this project are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Fixed

- Match the current SharePoint multipart format when exporting Loop pages, and discover the tenant SharePoint resource automatically when it was not captured during initial login.
- Send the registered Microsoft 365 client scenario required by the Loop Web Service for page listing, reads, creation, and updates.

## [0.2.1] - 2026-09-10

### Documentation

- Clarify direct npm execution, global installation, and local source-build configuration.

## [0.2.0] - 2026-09-10

### Added

- `loop_create_page` for creating Loop pages from Markdown.
- `loop_update_page` for appending, prepending, replacing a named section, replacing the complete body, and renaming pages.
- Loop Web Service token extraction and refresh, with an optional Azure CLI token fallback.

### Changed

- Rename the maintained fork to `ms-loop-mcp-server` and replace the upstream-oriented README with direct capability, limitation, local stdio, and SSH usage documentation.

### Fixed

- `MSLOOP_BROWSER=chromium` now selects Playwright's bundled browser instead of Chrome.
- Interactive login now respects its configured timeout and allows ten minutes for password/MFA flows.
- Bearer-token clients now reject unexpected hosts and insecure URLs before attaching credentials.
- SharePoint coordinates and Loop Web Service page ids now reject non-SharePoint hosts, and URL path segments are encoded.
- Substrate pagination ignores links that leave its expected HTTPS origin.
- Browser cookie import now uses exact-domain matching and private temporary files, and avoids shell interpolation when invoking platform credential tools.

### Documentation

- Document local clone-and-build usage for running the maintained fork without installing the upstream npm package.

## [0.1.2] - 2026-06-12

### Changed

- Update the package author contact details.

## [0.1.1] - 2026-06-12

### Changed

- Exclude compiled test files from the published package (smaller tarball, cleaner `dist`).

### Internal

- First release published via GitHub Actions using npm Trusted Publishing (OIDC), no stored token.

## [0.1.0] - 2026-06-12

Initial release.

### Added

- Browser-session auth that reuses the Loop web app's own first party client ID, no Azure app registration required. Tokens (Substrate, SharePoint, Graph) are extracted from MSAL local and session storage, cached AES-256-GCM encrypted in `~/.msloop-mcp-server/`, and refreshed over HTTP with a headless browser fallback.
- `loop_login`, `loop_status`, `loop_logout` auth tools.
- `loop_list_workspaces` — list accessible workspaces including the personal "My workspace".
- `loop_list_pages` — list pages in a workspace.
- `loop_get_page` — read a page's content as Markdown or HTML (via SharePoint's on-demand HTML export of the Fluid document).
- `loop_search` — search Loop files by keyword via Microsoft Graph.
- `loop_create_workspace` — experimental shared workspace creation.
- Cross platform SSO cookie import (macOS Keychain, Linux libsecret, Windows DPAPI) for instant silent first login.

[Unreleased]: https://github.com/vilsonrodrigues/ms-loop-mcp-server/compare/v0.2.1...HEAD
[0.2.1]: https://github.com/vilsonrodrigues/ms-loop-mcp-server/compare/v0.2.0...v0.2.1
[0.2.0]: https://github.com/vilsonrodrigues/ms-loop-mcp-server/compare/v0.1.2...v0.2.0
[0.1.2]: https://github.com/shayanline/msloop-mcp/compare/v0.1.1...v0.1.2
[0.1.1]: https://github.com/shayanline/msloop-mcp/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/shayanline/msloop-mcp/releases/tag/v0.1.0
