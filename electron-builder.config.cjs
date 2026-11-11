// Packaging, signing and update settings for electron-builder.
// Signing and notarizing switch on by themselves when the secrets below are in the
// environment (see SIGNING.md); without them you still get working, unsigned builds.
const pkg = require('./package.json');

// Where updates are published: GITHUB_REPOSITORY in GitHub Actions, otherwise
// KERTAS_UPDATE_REPO ("owner/repo") or the "repository" field of package.json.
const repoSlug = process.env.KERTAS_UPDATE_REPO || process.env.GITHUB_REPOSITORY
  || String(pkg.repository?.url || pkg.repository || '').replace(/^git\+/, '').replace(/^https:\/\/github\.com\//, '').replace(/\.git$/, '').replace(/^github:/, '');
// (the placeholder in package.json counts as not set)
const [owner, repo] = /^[\w.-]+\/[\w.-]+$/.test(repoSlug) && !repoSlug.startsWith('YOUR-GITHUB-NAME/') ? repoSlug.split('/') : [];

// On Windows `name` becomes the registry ProgID, so it is KertasOffice.<ext> there (the
// installer's Default apps registration in build/installer.nsh points at these); macOS
// shows `name` in Finder, so it stays readable there.
const WIN = process.platform === 'win32';
const assoc = (ext, name, role = 'Editor') => ({
  ext, role,
  name: WIN ? `KertasOffice.${ext}` : name,
  description: name,
  rank: 'Default', // macOS: Kertas is a full editor for these, not just a viewer
  icon: `build/filetypes/${ext}`, // .icns on macOS, .ico on Windows
  ...(ext === 'pdf' ? { mimeType: 'application/pdf' } : {}),
});

module.exports = {
  appId: 'my.kertas.office',
  productName: 'Kertas Office',
  copyright: `Copyright © ${new Date().getFullYear()} Kertas Office contributors`,
  directories: { output: 'release', buildResources: 'build' },
  files: ['dist/**', 'electron/**', 'package.json'],
  // Renderer code is bundled by Vite into dist/, so only the main process's own
  // dependencies (electron-updater, word-extractor) are shipped as node_modules.
  asar: true,
  icon: 'build/icon',
  fileAssociations: [
    assoc('docx', 'Word Document'),
    assoc('doc', 'Word 97-2003 Document'),
    assoc('rtf', 'Rich Text Document'),
    assoc('odt', 'OpenDocument Text'),
    assoc('xlsx', 'Excel Workbook'),
    assoc('xlsm', 'Excel Macro-Enabled Workbook'),
    assoc('xls', 'Excel 97-2003 Workbook'),
    assoc('ods', 'OpenDocument Spreadsheet'),
    assoc('csv', 'CSV File'),
    assoc('pdf', 'PDF Document'),
  ],
  // KERTAS_UPDATE_URL publishes to your own web server instead (any folder served over https)
  ...(process.env.KERTAS_UPDATE_URL
    ? { publish: [{ provider: 'generic', url: process.env.KERTAS_UPDATE_URL }] }
    : owner ? { publish: [{ provider: 'github', owner, repo, releaseType: 'draft' }] } : {}),

  win: {
    target: [{ target: 'nsis', arch: ['x64', 'arm64'] }],
    icon: 'build/icon.ico',
    // Signed when WIN_CSC_LINK / CSC_LINK (+ password) are set, or with Azure
    // Trusted Signing when the AZURE_* variables below are set.
    ...(process.env.AZURE_TRUSTED_SIGNING_ACCOUNT ? {
      azureSignOptions: {
        publisherName: process.env.AZURE_PUBLISHER_NAME || 'Kertas Office',
        endpoint: process.env.AZURE_TRUSTED_SIGNING_ENDPOINT,
        codeSigningAccountName: process.env.AZURE_TRUSTED_SIGNING_ACCOUNT,
        certificateProfileName: process.env.AZURE_TRUSTED_SIGNING_PROFILE,
      },
    } : {}),
  },
  nsis: {
    // File names carry no version, so a website can link to .../releases/latest/download/<name>
    artifactName: 'Kertas-Office-Setup.${ext}',
    oneClick: false,
    perMachine: false,
    allowToChangeInstallationDirectory: true,
    createDesktopShortcut: true,
    installerIcon: 'build/icon.ico',
    uninstallerIcon: 'build/icon.ico',
    license: 'LICENSE',
    include: 'build/installer.nsh', // registers Kertas under Settings > Default apps
  },

  mac: {
    target: [{ target: 'dmg', arch: ['arm64', 'x64'] }, { target: 'zip', arch: ['arm64', 'x64'] }], // zip is what auto-update downloads
    artifactName: 'Kertas-Office-mac-${arch}.${ext}',
    category: 'public.app-category.productivity',
    icon: 'build/icon.icns',
    hardenedRuntime: true,
    gatekeeperAssess: false,
    entitlements: 'build/entitlements.mac.plist',
    entitlementsInherit: 'build/entitlements.mac.plist',
    // Signed with CSC_LINK (Developer ID Application .p12) and notarized when either an
    // App Store Connect API key (APPLE_API_KEY = path to the .p8, APPLE_API_KEY_ID,
    // APPLE_API_ISSUER) or an Apple ID (APPLE_ID, APPLE_APP_SPECIFIC_PASSWORD, APPLE_TEAM_ID) is set.
    notarize: !!((process.env.APPLE_API_KEY && process.env.APPLE_API_KEY_ID && process.env.APPLE_API_ISSUER)
      || (process.env.APPLE_ID && process.env.APPLE_APP_SPECIFIC_PASSWORD && process.env.APPLE_TEAM_ID)),
  },
  dmg: { sign: false },

  linux: {
    target: ['AppImage', 'deb'],
    category: 'Office',
    icon: 'build/icons',
    syncDesktopName: true,
    synopsis: 'Documents, spreadsheets and PDFs',
    // Debian packages need a maintainer with an email address; set KERTAS_MAINTAINER for releases
    maintainer: process.env.KERTAS_MAINTAINER || 'Kertas Office contributors <maintainers@kertas.invalid>',
    mimeTypes: [
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/msword', 'application/rtf',
      'application/vnd.oasis.opendocument.text', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-excel', 'application/vnd.oasis.opendocument.spreadsheet', 'text/csv', 'application/pdf',
    ],
  },
  appImage: { artifactName: 'Kertas-Office-linux-${arch}.${ext}' },
  deb: { artifactName: 'kertas-office-${arch}.${ext}', depends: ['libnotify4', 'libxtst6', 'libnss3'], recommends: ['libreoffice-core'] },
};
