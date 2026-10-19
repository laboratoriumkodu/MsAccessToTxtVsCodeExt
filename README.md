# Access Source Exporter

Access Source Exporter brings the Access database source-export workflow into
Visual Studio Code. Run **Access: Export Source** from the Command Palette,
choose an `.accdb` or `.mdb` database and a destination folder, and the
extension runs the bundled PowerShell exporter for you.

The exporter writes the database's VBA modules, forms, reports, macros, query
SQL, table schema, and relationships as text files in organized folders. This
makes Access project source easier to inspect, compare in version control, and
maintain alongside application code. Database records are not exported.

The extension packages the PowerShell script within this project and invokes
it directly; it does not reimplement the export logic. The extension can
therefore be built and packaged independently of the original repository, and
the PowerShell script remains independently runnable.

## Using the extension

1. Open the Command Palette in VS Code.
2. Run **Access: Export Source**.
3. Select an Access database and an output folder.
4. Follow progress in the notification and view exporter output in the
   **Access Source Export** Output channel.

Exports can be cancelled while running.

## Requirements

- Windows
- Windows PowerShell 5.1 (`powershell.exe`)
- Microsoft Access installed and activated

The exporter uses Access COM automation. PowerShell's process architecture
must be compatible with the installed Access architecture.

## Development

Open this folder in VS Code and start the **Run Extension** launch
configuration to build the extension and open an Extension Development Host.
The launch task runs `npm run compile` before starting. Use **Access: Export
Source** in the host window to try the command.

Run `npm test` to compile and run the extension tests in a temporary VS Code
instance.
