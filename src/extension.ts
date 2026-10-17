import * as vscode from 'vscode';
import { spawn } from 'node:child_process';

function buildPowerShellArguments(
	scriptPath: string,
	databasePath: string,
	outputPath: string
): string[] {
	return [
		'-NoProfile',
		'-ExecutionPolicy',
		'Bypass',
		'-File',
		scriptPath,
		'-AccdbPath',
		databasePath,
		'-OutputPath',
		outputPath,
	];
}

async function selectDatabase(): Promise<vscode.Uri | undefined> {
	const selections = await vscode.window.showOpenDialog({
		canSelectFiles: true,
		canSelectFolders: false,
		canSelectMany: false,
		openLabel: 'Select Access database',
		filters: {
			'Access databases': ['accdb', 'mdb'],
		},
	});

	return selections?.[0];
}

async function selectOutputFolder(): Promise<vscode.Uri | undefined> {
	const selections = await vscode.window.showOpenDialog({
		canSelectFiles: false,
		canSelectFolders: true,
		canSelectMany: false,
		openLabel: 'Select output folder',
	});

	return selections?.[0];
}

type ExportResult = 'completed' | 'cancelled';

function runExporter(
	scriptPath: string,
	databasePath: string,
	outputPath: string,
	token: vscode.CancellationToken,
	output: vscode.OutputChannel
): Promise<ExportResult> {
	return new Promise((resolve, reject) => {
		if (token.isCancellationRequested) {
			resolve('cancelled');
			return;
		}

		const child = spawn(
			'powershell.exe',
			buildPowerShellArguments(scriptPath, databasePath, outputPath),
			{ windowsHide: true }
		);

		let settled = false;
		let stderr = '';
		const cancellation = token.onCancellationRequested(() => {
			output.appendLine('Cancellation requested; terminating PowerShell.');
			if (!child.kill()) {
				output.appendLine('PowerShell had already exited before it could be terminated.');
			}
		});

		const rejectOnce = (error: Error): void => {
			if (settled) {
				return;
			}
			settled = true;
			cancellation.dispose();
			reject(error);
		};

		child.stdout.setEncoding('utf8');
		child.stderr.setEncoding('utf8');
		child.stdout.on('data', (chunk: string) => output.append(chunk));
		child.stderr.on('data', (chunk: string) => {
			stderr = (stderr + chunk).slice(-4000);
			output.append(chunk);
		});

		child.once('error', (error: NodeJS.ErrnoException) => {
			if (error.code === 'ENOENT') {
				rejectOnce(
					new Error(
						'Windows PowerShell (powershell.exe) was not found. Install or repair Windows PowerShell 5.1 and ensure it is available on PATH.'
					)
				);
				return;
			}
			rejectOnce(error);
		});
		child.once('close', (code, signal) => {
			if (settled) {
				return;
			}
			settled = true;
			cancellation.dispose();

			if (token.isCancellationRequested) {
				resolve('cancelled');
				return;
			}
			if (code === 0) {
				resolve('completed');
				return;
			}

			reject(new Error(
				signal
					? `PowerShell was terminated by signal ${signal}.${stderr ? `\n${stderr.trim()}` : ''}`
					: `The exporter exited with code ${code ?? 'unknown'}.${stderr ? `\n${stderr.trim()}` : ''}`
			));
		});
	});
}

function getActionableError(error: unknown): string {
	const message = error instanceof Error ? error.message : String(error);
	if (/COMException|COM class factory|Access\.Application|class factory|class not registered|0x80040154|OpenCurrentDatabase/i.test(message)) {
		return `Microsoft Access could not be opened or automated. Confirm that Access is installed and activated, and that its installation matches the PowerShell process architecture. Details: ${message}`;
	}
	return message;
}

async function exportSource(
	context: vscode.ExtensionContext,
	output: vscode.OutputChannel
): Promise<void> {
	if (process.platform !== 'win32') {
		output.appendLine('Export failed: unsupported operating system.');
		output.show(true);
		void vscode.window.showErrorMessage(
			'Access source export is supported only on Windows because it requires Windows PowerShell and Microsoft Access COM automation.'
		);
		return;
	}

	try {
		const database = await selectDatabase();
		if (!database) {
			return;
		}

		const outputFolder = await selectOutputFolder();
		if (!outputFolder) {
			return;
		}

		const scriptPath = vscode.Uri.joinPath(
			context.extensionUri,
			'dist',
			'Export-AccessSource.ps1'
		).fsPath;

		await vscode.window.withProgress(
			{
				location: vscode.ProgressLocation.Notification,
				title: 'Exporting Access source',
				cancellable: true,
			},
			async (progress, token) => {
				output.clear();
				output.appendLine(`Database: ${database.fsPath}`);
				output.appendLine(`Output folder: ${outputFolder.fsPath}`);
				output.show(true);

				progress.report({ message: 'Running PowerShell exporter...' });
				const result = await runExporter(
					scriptPath,
					database.fsPath,
					outputFolder.fsPath,
					token,
					output
				);

				if (result === 'cancelled') {
					output.appendLine('Export cancelled.');
					void vscode.window.showWarningMessage(
						'Access source export was cancelled.'
					);
					return;
				}

				output.appendLine('Export completed successfully.');
				void vscode.window.showInformationMessage(
					'Access source export completed.'
				);
			}
		);
	} catch (error) {
		const message = getActionableError(error);
		output.appendLine(`Export failed: ${message}`);
		output.show(true);
		void vscode.window.showErrorMessage(`Access source export failed: ${message}`);
	}
}

export function activate(context: vscode.ExtensionContext): void {
	const output = vscode.window.createOutputChannel('Access Source Export');
	const command = vscode.commands.registerCommand(
		'msaccesstotxt.exportSource',
		() => exportSource(context, output)
	);

	context.subscriptions.push(command);
	context.subscriptions.push(output);
}

export function deactivate() {}
