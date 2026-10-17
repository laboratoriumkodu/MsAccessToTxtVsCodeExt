import * as vscode from 'vscode';

export function activate(context: vscode.ExtensionContext): void {
	const command = vscode.commands.registerCommand('msaccesstotxt.exportSource', () => {
		void vscode.window.showInformationMessage(
			'Access source export is not implemented yet.'
		);
	});

	context.subscriptions.push(command);
}

export function deactivate() {}
