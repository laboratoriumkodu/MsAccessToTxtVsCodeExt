# Publishing the Extension

## 1. One-time account setup

- Create a publisher at the [VS Code Marketplace management page](https://marketplace.visualstudio.com/manage) (sign in with a Microsoft/Azure DevOps account).
- Generate an Azure DevOps **Personal Access Token** (Organization: "All accessible organizations", scope: **Marketplace → Manage**).

## 2. Add required manifest fields

`package.json` needs these fields before publishing:

- `publisher` — your publisher ID from step 1
- `repository` — link to this repo
- `license` — e.g. `"SEE LICENSE IN ../LICENSE"` or an SPDX id
- `icon` — path to a 128×128 PNG
- `categories`/`keywords` — already has categories, could add keywords

## 3. Install the publishing tool

```powershell
npm install --save-dev @vscode/vsce
```

## 4. Package and smoke-test locally

```powershell
npx vsce package
code --install-extension msaccesstotxt-0.0.1.vsix
```

Verify the PowerShell script is actually bundled (check `dist/` and the vsix contents) before publishing.

## 5. Login and publish

```powershell
npx vsce login <publisher-id>
npx vsce publish
```

`vsce publish patch|minor|major` will also bump the version and tag it.
