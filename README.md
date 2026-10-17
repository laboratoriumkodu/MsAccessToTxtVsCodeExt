# Access Source Exporter

This VS Code extension will provide a command to export source files from
Microsoft Access databases by invoking the existing PowerShell exporter.
The standalone script remains at `../script/Export-AccessSource.ps1`.

## Current status

The extension scaffold and **Access: Export Source** command are in place.
The export workflow is not implemented yet.

## Requirements

The completed extension is intended for Windows with PowerShell and Microsoft
Access installed. The exporter uses Access COM automation.

## Development

Open this folder in VS Code and run the `Run Extension` launch configuration to
start an Extension Development Host. Run `npm run compile` to build and check
the extension.
