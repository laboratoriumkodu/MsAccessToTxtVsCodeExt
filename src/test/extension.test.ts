import * as assert from 'assert';
import { EventEmitter } from 'node:events';
import * as vscode from 'vscode';
import {
	activate,
	buildPowerShellArguments,
	exportSource,
	getActionableError,
	runExporter,
} from '../extension';

suite('Access Source Exporter', () => {
	const originalShowOpenDialog = vscode.window.showOpenDialog;
	const originalWithProgress = vscode.window.withProgress;
	const originalShowInformationMessage = vscode.window.showInformationMessage;
	const originalShowWarningMessage = vscode.window.showWarningMessage;
	const originalShowErrorMessage = vscode.window.showErrorMessage;

	teardown(() => {
		vscode.window.showOpenDialog = originalShowOpenDialog;
		vscode.window.withProgress = originalWithProgress;
		vscode.window.showInformationMessage = originalShowInformationMessage;
		vscode.window.showWarningMessage = originalShowWarningMessage;
		vscode.window.showErrorMessage = originalShowErrorMessage;
	});

	test('buildPowerShellArguments keeps path spacing intact', () => {
		const args = buildPowerShellArguments(
			'C:\\Program Files\\Access Export\\Export-AccessSource.ps1',
			'C:\\Users\\Example\\My Access Database.accdb',
			'C:\\Users\\Example\\Exported Source Folder'
		);

		assert.deepStrictEqual(args, [
			'-NoProfile',
			'-ExecutionPolicy',
			'Bypass',
			'-File',
			'C:\\Program Files\\Access Export\\Export-AccessSource.ps1',
			'-AccdbPath',
			'C:\\Users\\Example\\My Access Database.accdb',
			'-OutputPath',
			'C:\\Users\\Example\\Exported Source Folder',
		]);
	});

	test('getActionableError explains Access automation failures', () => {
		const message = getActionableError(
			'COMException: Microsoft Access could not open the database. Access.Application class factory is not registered.'
		);

		assert.match(
			message,
			/Microsoft Access could not be opened or automated/i
		);
		assert.match(message, /Access is installed/i);
	});

	test('runExporter resolves cancelled when the operation is cancelled', async () => {
		const output = vscode.window.createOutputChannel('test-output');
		const token = {
			isCancellationRequested: true,
			onCancellationRequested: (_listener: () => void) => ({ dispose: () => undefined }),
		} as vscode.CancellationToken;

		const result = await runExporter(
			'C:\\temp\\Export-AccessSource.ps1',
			'C:\\temp\\Database.accdb',
			'C:\\temp\\Out',
			token,
			output
		);

		assert.strictEqual(result, 'cancelled');
		output.dispose();
	});

	test('exportSource completes the command flow with selected database and folder', async () => {
		const context = {
			extensionUri: vscode.Uri.file('C:\\work\\msaccesstotxt\\vscode-extension'),
			subscriptions: [],
		} as unknown as vscode.ExtensionContext;
		const output = vscode.window.createOutputChannel('Access Source Export');
		const databasePath = 'C:\\Users\\Test\\My Access Database.accdb';
		const outputPath = 'C:\\Users\\Test\\Exported Access Source';
		const dialogResults = [
			[vscode.Uri.file(databasePath)],
			[vscode.Uri.file(outputPath)],
		];
		let spawnedArgs: string[] | undefined;
		const infoMessages: string[] = [];
		const token = {
			isCancellationRequested: false,
			onCancellationRequested: (_listener: () => void) => ({ dispose: () => undefined }),
		} as vscode.CancellationToken;

		(vscode.window as typeof vscode.window & { showOpenDialog: typeof vscode.window.showOpenDialog }).showOpenDialog = async () => {
			return dialogResults.shift() ?? undefined;
		};
		(vscode.window as unknown as {
			withProgress: <T>(
				options: unknown,
				task: (
					progress: { report: (message: unknown) => void },
					token: vscode.CancellationToken
				) => Thenable<T>
			) => Thenable<T>;
		}).withProgress = async <T>(
			_options: unknown,
			task: (
				progress: { report: (message: unknown) => void },
				token: vscode.CancellationToken
			) => Thenable<T>
		): Promise<T> => {
			return task({ report: () => undefined }, token) as Promise<T>;
		};
		(vscode.window as typeof vscode.window & { showInformationMessage: typeof vscode.window.showInformationMessage }).showInformationMessage = async (message: string) => {
			infoMessages.push(message);
			return undefined;
		};
		(vscode.window as typeof vscode.window & { showWarningMessage: typeof vscode.window.showWarningMessage }).showWarningMessage = async () => undefined;
		(vscode.window as typeof vscode.window & { showErrorMessage: typeof vscode.window.showErrorMessage }).showErrorMessage = async () => undefined;

		const child = new EventEmitter() as EventEmitter & {
			stdout: EventEmitter & { setEncoding: () => void };
			stderr: EventEmitter & { setEncoding: () => void };
			kill: () => boolean;
		};
		child.stdout = new EventEmitter() as EventEmitter & { setEncoding: () => void };
		child.stderr = new EventEmitter() as EventEmitter & { setEncoding: () => void };
		child.stdout.setEncoding = () => undefined;
		child.stderr.setEncoding = () => undefined;
		child.kill = () => true;
		child.once = ((event: string, listener: (...args: unknown[]) => void) => {
			if (event === 'close') {
				setTimeout(() => listener(0, null), 0);
			}
			return child;
		}) as typeof child.once;
		const spawnStub = (_command: string, args: readonly string[]) => {
			spawnedArgs = [...args];
			return child as any;
		};

		activate(context);
		await exportSource(context, output, spawnStub);

		assert.ok(spawnedArgs);
		assert.deepStrictEqual(spawnedArgs, [
			'-NoProfile',
			'-ExecutionPolicy',
			'Bypass',
			'-File',
			'c:\\work\\msaccesstotxt\\vscode-extension\\dist\\Export-AccessSource.ps1',
			'-AccdbPath',
			'c:\\Users\\Test\\My Access Database.accdb',
			'-OutputPath',
			'c:\\Users\\Test\\Exported Access Source',
		]);
		assert.ok(infoMessages.includes('Access source export completed.'));

		for (const subscription of context.subscriptions) {
			subscription.dispose?.();
		}
		output.dispose();
	});

	test('cancelled dialog selection aborts export without spawning PowerShell', async () => {
		const context = {
			extensionUri: vscode.Uri.file('C:\\work\\msaccesstotxt\\vscode-extension'),
			subscriptions: [],
		} as unknown as vscode.ExtensionContext;
		const output = vscode.window.createOutputChannel('Access Source Export');
		let spawnCalled = false;

		(vscode.window as typeof vscode.window & { showOpenDialog: typeof vscode.window.showOpenDialog }).showOpenDialog = async () => undefined;
		(vscode.window as any).withProgress = async () => undefined;
		(vscode.window as typeof vscode.window & { showErrorMessage: typeof vscode.window.showErrorMessage }).showErrorMessage = async () => undefined;
		const spawnStub = () => {
			spawnCalled = true;
			throw new Error('spawn should not be called when selection is cancelled');
		};

		await exportSource(context, output, spawnStub);

		assert.strictEqual(spawnCalled, false);
		output.dispose();
	});
});
