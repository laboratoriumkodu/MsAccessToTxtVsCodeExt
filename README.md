# Access Source Exporter

Use **Access: Export Source** from the Command Palette to select an Access
database and an output folder. The extension invokes the existing PowerShell
exporter. Its copy in `script/` is included in this project, so the extension
can be built and packaged independently of the standalone exporter in the
parent repository.

## Current status

The command flow is implemented. The build copies the bundled exporter into
the extension package, so an installed extension does not depend on the
repository being present. Export progress can be cancelled, and PowerShell
output and errors appear in the **Access Source Export** Output channel.

## Requirements

The extension requires Windows, Windows PowerShell 5.1 (`powershell.exe`), and
Microsoft Access installed and activated. The exporter uses Access COM
automation and its process architecture must be compatible with the Access
installation.

## Development

Open this folder in VS Code and start the `Run Extension` launch configuration
to build the extension and open an Extension Development Host. The launch task
runs `npm run compile` before each start. Use the **Access: Export Source**
command in the host window to try the extension.

Run `npm test` separately to execute the extension tests; the test runner opens
and closes its own temporary VS Code instance.
