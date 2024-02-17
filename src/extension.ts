// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from 'vscode';
import * as XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';

// This method is called when your extension is activated
// Your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {

	// Use the console to output diagnostic information (console.log) and errors (console.error)
	// This line of code will only be executed once when your extension is activated
	console.log('Congratulations, your extension "excelworld" is now active!');

	// The command has been defined in the package.json file
	// Now provide the implementation of the command with registerCommand
	// The commandId parameter must match the command field in package.json
	let disposable = vscode.commands.registerCommand('extension.readExcel', () => {
		// The code you place here will be executed every time your command is executed
		// Display a message box to the user
		// vscode.window.showInformationMessage('Hello World from excelWorld varan!');
		vscode.window.showOpenDialog({
			canSelectMany: false,
			filters: {
				'Excel files': ['xlsx']
			}
		}).then(fileUri => {
			if (fileUri && fileUri[0]) {
				let workbook = XLSX.readFile(fileUri[0].fsPath);
				let sheet_name_list = workbook.SheetNames;
				sheet_name_list.forEach(function(y){
					var worksheet = workbook.Sheets[y];
					var excelDataArray = XLSX.utils.sheet_to_json(worksheet, {header: 1});

					var excelData : { [key: string]: any; } = {};
					excelDataArray.forEach((row:any) => {
						if(row.length >= 2) {
							excelData[row[0]]=row[1];
						}
					});

					//Get the active editor
					let activeEditor = vscode.window.activeTextEditor;
					if (activeEditor && activeEditor.document.languageId === 'json') {
						let jsonPath = activeEditor.document.uri.fsPath;
						let jsonData = JSON.parse(fs.readFileSync(jsonPath,'utf-8'));

						//Compare JSON data with Excel data
						for(let key in excelData) {
							if(excelData.hasOwnProperty(key)){
								if (jsonData.hasOwnProperty(key) && jsonData[key] === excelData[key]) {
                                } else {
                                    // vscode.window.showInformationMessage(`No match found for: ${key} - ${excelData[key]}`);
									if (activeEditor) {
										const text = activeEditor.document.getText();
										const position = activeEditor.document.positionAt(text.lastIndexOf('}'));
										const workspaceEdit = new vscode.WorkspaceEdit();
										workspaceEdit.insert(activeEditor.document.uri, position, `,\n"${key}": "${excelData[key]}"`);
										vscode.workspace.applyEdit(workspaceEdit);
									}
                                }
							}
						}
					}
					else {
						vscode.window.showErrorMessage('Please open a JSON file');
					}
				});
			}
		});
	});

	context.subscriptions.push(disposable);
}

// This method is called when your extension is deactivated
export function deactivate() {}
