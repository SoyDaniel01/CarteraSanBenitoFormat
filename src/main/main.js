const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const log = require('electron-log');

log.transports.file.level = 'info';
log.catchErrors({ showDialog: false });

let mainWindow;

const isMac = process.platform === 'darwin';

const projectRoot = path.join(__dirname, '..', '..');

/**
 * Attempt to resolve a Python interpreter from the local virtual environment.
 */
function resolveDevPython() {
  const venvDir = path.join(projectRoot, '.venv');
  const binariesDir = process.platform === 'win32'
    ? path.join(venvDir, 'Scripts')
    : path.join(venvDir, 'bin');

  const candidates = process.platform === 'win32'
    ? [
        path.join(binariesDir, 'python.exe'),
        path.join(binariesDir, 'python'),
      ]
    : [
        path.join(binariesDir, 'python3'),
        path.join(binariesDir, 'python'),
      ];

  return candidates.find(candidate => fs.existsSync(candidate));
}

/**
 * Resolve the absolute path to the Python processor bundled with the app.
 * - In desarrollo: usa `.venv` si está disponible, o python del sistema.
 * - En producción: prefiere un ejecutable empaquetado (PyInstaller).
 */
function getPythonInvocation() {
  const scriptPath = path.join(projectRoot, 'python', 'processor.py');

  if (!app.isPackaged) {
    const localPython = resolveDevPython();
    if (localPython) {
      return { command: localPython, args: [scriptPath] };
    }

    const pythonCmd = process.platform === 'win32' ? 'python' : 'python3';
    return { command: pythonCmd, args: [scriptPath] };
  }

  const resourcesPath = process.resourcesPath;
  const packagedDir = path.join(resourcesPath, 'python');
  const executableName = process.platform === 'win32' ? 'processor.exe' : 'processor';

  const executableCandidates = [
    path.join(packagedDir, executableName),
    path.join(packagedDir, 'processor', executableName),
    path.join(packagedDir, 'dist', executableName),
    path.join(packagedDir, 'dist', 'processor', executableName),
    path.join(packagedDir, 'bin', executableName)
  ];

  for (const candidate of executableCandidates) {
    if (fs.existsSync(candidate)) {
      return { command: candidate, args: [] };
    }
  }

  // Fallback to python runtime shipped inside extraResources
  const packagedScript = path.join(packagedDir, 'processor.py');
  const pythonCmd = process.platform === 'win32' ? 'python' : 'python3';
  return { command: pythonCmd, args: [packagedScript] };
}

function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: 900,
    height: 600,
    minWidth: 820,
    minHeight: 520,
    show: false,
    title: 'Cartera San Benito Format',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      sandbox: true,
      nodeIntegration: false
    }
  });

  mainWindow.loadFile(path.join(__dirname, '..', 'renderer', 'index.html'));

  mainWindow.webContents.on('did-finish-load', () => {
    sendToRenderer('log-message', { level: 'info', message: `Aplicación iniciada. Versión ${app.getVersion()}` });
  });

  mainWindow.on('ready-to-show', () => {
    mainWindow?.show();
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });

  mainWindow.removeMenu();

  return mainWindow;
}

function sendToRenderer(channel, payload) {
  if (mainWindow && mainWindow.webContents) {
    mainWindow.webContents.send(channel, payload);
  }
}

function runPythonProcessor(inputPath, outputPath) {
  return new Promise((resolve, reject) => {
    try {
      const { command, args } = getPythonInvocation();
      const allArgs = [...args, inputPath, outputPath];

      log.info(`Invocando Python: ${command} ${allArgs.join(' ')}`);
      sendToRenderer('log-message', { level: 'info', message: `Ejecutando procesamiento con Python (${path.basename(inputPath)})` });

      const pyProcess = spawn(command, allArgs, { stdio: ['ignore', 'pipe', 'pipe'] });

      pyProcess.stdout.on('data', data => {
        const message = data.toString().trim();
        if (message) {
          log.info(`Python stdout: ${message}`);
          sendToRenderer('log-message', { level: 'info', message });
        }
      });

      pyProcess.stderr.on('data', data => {
        const message = data.toString().trim();
        if (message) {
          log.error(`Python stderr: ${message}`);
          sendToRenderer('log-message', { level: 'error', message });
        }
      });

      pyProcess.on('error', error => {
        log.error('Error al ejecutar Python', error);
        reject(error);
      });

      pyProcess.on('close', code => {
        if (code === 0) {
          log.info(`Procesamiento finalizado. Archivo listo en ${outputPath}`);
          resolve(outputPath);
        } else {
          const err = new Error(`Python finalizó con código ${code}`);
          err.code = code;
          reject(err);
        }
      });
    } catch (err) {
      reject(err);
    }
  });
}

async function handleProcessRequest(filePath) {
  if (!filePath) {
    throw new Error('No se recibió la ruta del archivo.');
  }

  if (path.extname(filePath).toLowerCase() !== '.xlsx') {
    throw new Error('El archivo debe tener extensión .xlsx');
  }

  const destinationDir = path.dirname(filePath);
  const parsedName = path.parse(filePath).name;
  const outputFile = path.join(destinationDir, `${parsedName}_processed.xlsx`);

  sendToRenderer('processing-state', { status: 'running', input: filePath, output: outputFile });

  const processedPath = await runPythonProcessor(filePath, outputFile);

  sendToRenderer('processing-state', { status: 'completed', output: processedPath });
  return processedPath;
}

function registerIpcHandlers() {
  ipcMain.handle('dialog:select-file', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
      title: 'Seleccionar archivo Excel',
      filters: [{ name: 'Excel', extensions: ['xlsx'] }],
      properties: ['openFile']
    });

    if (result.canceled || result.filePaths.length === 0) {
      return undefined;
    }

    return result.filePaths[0];
  });

  ipcMain.handle('processor:run', async (_event, filePath) => {
    try {
      const outputPath = await handleProcessRequest(filePath);
      return { success: true, outputPath };
    } catch (error) {
      log.error('Fallo al procesar archivo', error);
      sendToRenderer('processing-state', { status: 'failed', error: error.message });
      return { success: false, error: error.message };
    }
  });

  ipcMain.handle('system:reveal-in-folder', (_event, targetPath) => {
    if (!targetPath) {
      return false;
    }

    try {
      shell.showItemInFolder(targetPath);
      return true;
    } catch (error) {
      log.error('No se pudo abrir la carpeta de salida', error);
      return false;
    }
  });

  ipcMain.on('log:forward', (_event, payload) => {
    if (payload?.message) {
      log.info(`[Renderer] ${payload.message}`);
    }
  });
}

function setupAutoUpdater(updater) {
  if (!updater) {
    log.warn('AutoUpdater no disponible en este entorno.');
    return;
  }

  updater.logger = log;
  updater.logger.transports.file.level = 'info';

  updater.on('update-available', () => {
    log.info('Actualización disponible.');
    sendToRenderer('update-event', { type: 'available' });
  });

  updater.on('update-not-available', () => {
    log.info('No hay actualizaciones disponibles.');
    sendToRenderer('update-event', { type: 'not-available' });
  });

  updater.on('error', error => {
    log.error('Error en autoUpdater', error);
    sendToRenderer('update-event', { type: 'error', message: error?.message ?? String(error) });
  });

  updater.on('download-progress', progressObj => {
    sendToRenderer('update-event', { type: 'progress', progress: progressObj });
  });

  updater.on('update-downloaded', () => {
    log.info('Actualización descargada, lista para instalar.');
    sendToRenderer('update-event', { type: 'downloaded' });
  });
}

app.whenReady().then(() => {
  createMainWindow();
  registerIpcHandlers();

  log.info(`Aplicación inicializada. Versión ${app.getVersion()}`);

  try {
    const { autoUpdater } = require('electron-updater');
    setupAutoUpdater(autoUpdater);
    autoUpdater.checkForUpdatesAndNotify().catch(error => {
      log.error('Fallo al buscar actualizaciones', error);
    });
  } catch (error) {
    log.warn('AutoUpdater deshabilitado o no disponible en modo desarrollo.', error);
  }

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createMainWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (!isMac) {
    app.quit();
  }
});
