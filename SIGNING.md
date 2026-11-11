# Releasing, signing and updates

Kertas Office ships through GitHub Releases. Work lands on `develop`; merging
`develop` into `main` builds every
installer in GitHub Actions (`.github/workflows/release.yml`), signs it when the
certificates are configured, and uploads it to a **draft** release. Publishing the
draft rolls the update out: installed copies check every 6 hours (and 8 seconds
after start), download in the background, and show **Restart to update** in the
title bar. Unsaved work is offered for saving before the restart.

## Cutting a release

1. On `develop`, bump `version` in `package.json` (for example `1.0.1`) and merge that change.
2. Open a pull request from `develop` to `main` and merge it. This starts the Release workflow.
3. Approve the `release` environment when asked, wait for the builds, check the draft
   release `v1.0.1`, add notes and press **Publish release**. Publishing creates the tag.

Without any secrets the workflow still produces working installers; they are just
unsigned (Windows SmartScreen and macOS Gatekeeper will warn).

## Code signing

Signing secrets live in the `release` environment (Settings > Environments), which only runs on `main` can use.

### Windows (pick one)

| Secret | What |
|---|---|
| `WIN_CERT_PFX_BASE64` | Code-signing certificate exported as .pfx, then `base64 -w0 cert.pfx` |
| `WIN_CERT_PASSWORD` | Its password |

Or **Azure Trusted Signing** (cheapest way to a trusted Windows signature, about USD 10 a month):
`AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`, `AZURE_TRUSTED_SIGNING_ENDPOINT`
(for example `https://eus.codesigning.azure.net`), `AZURE_TRUSTED_SIGNING_ACCOUNT`,
`AZURE_TRUSTED_SIGNING_PROFILE`, `AZURE_PUBLISHER_NAME` (must match the certificate subject).

Hardware-token (EV) certificates cannot be used from CI this way; sign those builds on
your own machine with `npm run release` and the token plugged in.

### macOS

Needs an Apple Developer account. Add these to the `release` environment
(Settings > Environments > release > Environment secrets):

| Secret | What |
|---|---|
| `MAC_CERT_P12_BASE64` | "Developer ID Application" certificate exported from Keychain as .p12, base64 |
| `MAC_CERT_PASSWORD` | Its password |
| `APPLE_API_KEY_P8_BASE64` | App Store Connect API key (.p8), base64. App Store Connect > Users and Access > Integrations |
| `APPLE_API_KEY_ID` | The key's ID |
| `APPLE_API_ISSUER` | The Issuer ID shown above the key list |

The app is built with the hardened runtime (`build/entitlements.mac.plist`) and
notarized automatically when the API key secrets are present. macOS auto-update
only works for signed builds.

### Linux

AppImage and .deb are not signed. Set the repository variable `KERTAS_MAINTAINER`
to `Kertas Office <hello@kertasoffice.com>` so the .deb lists a real maintainer.

## Updates from your own server instead of GitHub

Build with `KERTAS_UPDATE_URL=https://downloads.example.com/kertas/ npm run dist` and
upload everything in `release/` (installers plus the `latest*.yml` files) to that
folder. Installed copies read `latest.yml` / `latest-mac.yml` / `latest-linux.yml` from there.

## Turning updates off

Users can untick **Download updates automatically** in About (Ctrl+K > About).
Administrators can set the environment variable `KERTAS_NO_UPDATE=1`.

## Building locally

```bash
npm run dist        # installers for this OS in release/, no upload
npm run icons       # re-render build/ icons after editing build/icon-src/*.svg
```
