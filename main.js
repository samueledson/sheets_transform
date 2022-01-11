var electron = require('electron');
var XLSX = require('xlsx');
var app = electron.app;

var win = null;

function createWindow() {
	if (win) return;
	win = new electron.BrowserWindow({
		width: 800, height: 600,
		webPreferences: {
			worldSafeExecuteJavaScript: true, // required for Electron 12+
			contextIsolation: false, // required for Electron 12+
			nodeIntegration: true,
			enableRemoteModule: true
		},
		//icon: "icon.ico",
		autoHideMenuBar: true
	});
	win.loadURL("file://" + __dirname + "/index.html");
	//win.webContents.openDevTools();
	win.on('closed', function () { win = null; });
}
if (app.setAboutPanelOptions) app.setAboutPanelOptions({ applicationName: 'transformar-planilhas', applicationVersion: "0.1", copyright: "(C) 2022 Samuel Edson" });
//app.on('open-file', function () { console.log(arguments); });
app.on('ready', createWindow);
app.on('activate', createWindow);
app.on('window-all-closed', function () { if (process.platform !== 'darwin') app.quit(); });