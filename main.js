const { app, BrowserWindow, net, ipcMain } = require('electron');
const path = require('path');
const isDev = require('electron-is-dev');

//const python = require('child_process').spawn('python', ['./src/server/py/PythonAPI.py']);
const { execFile, spawn } = require('child_process');
const pythonLangPath = path.join(__dirname, "src", "server", "py", "venv", "Scripts", "python.exe");
const pythonPath = isDev 
  ? path.join(__dirname, "src","server","py", "PythonAPI.py")
  : path.join(process.resourcesPath, "python","PythonAPI.exe");
console.log("python lang path: ", pythonLangPath)
console.log("python path: ", pythonPath)
let pythonAPI = !isDev && execFile(pythonPath, ['--version'], (error, stdout, stderr) => {
      if (error) {
        throw error;
      }
      console.log(stdout);
    });

function createWindow () {

  const win = new BrowserWindow({
    width: 800,
    height: 600,
    titleBarStyle: 'hidden',
    frame: false,
    webPreferences: {
      nodeIntegration: true,
      enableRemoteModule: true
    }
  })

  win.loadFile('index.html');
  win.maximize();
  win.webContents.openDevTools()
}

app.whenReady().then(createWindow)

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    
    !isDev && pythonAPI.kill();
    pythonAPI = null;
    app.quit()
  }
})

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow()
  }
})