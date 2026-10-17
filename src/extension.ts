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

function runExporter(
	scriptPath: string,
	databasePath: string,
	outputPath: string
): Promise<void> {
	return new Promise((resolve, reject) => {
		const child = spawn(
			'powershell.exe',
			buildPowerShellArguments(scriptPath, databasePath, outputPath),
			{ windowsHide: true }
		);

		child.stdout?.resume();
		child.stderr?.resume();

		child.once('error', reject);
		child.once('close', (code, signal) => {
			if (code === 0) {
				resolve();
				return;
			}

			reject(
				new Error(
					signal
						? `PowerShell was terminated by signal ${signal}.`
						: `The exporter exited with code ${code ?? 'unknown'}.`
				)
			);
		});
	});
}

async function exportSource(context: vscode.ExtensionContext): Promise<void> {
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

		await runExporter(scriptPath, database.fsPath, outputFolder.fsPath);
		void vscode.window.showInformationMessage(
			'Access source export completed.'
		);
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		void vscode.window.showErrorMessage(`Access source export failed: ${message}`);
	}
}

export function activate(context: vscode.ExtensionContext): void {
	const command = vscode.commands.registerCommand(
		'msaccesstotxt.exportSource',
		() => exportSource(context)
	);

	context.subscriptions.push(command);
}

export function deactivate() {}
