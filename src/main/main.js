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
  const platform = process.platform;

  if (!app.isPackaged) {
    log.info('Modo desarrollo: buscando Python local');
    const localPython = resolveDevPython();
    if (localPython) {
      log.info(`Usando Python local: ${localPython}`);
      return { command: localPython, args: [scriptPath] };
    }

    const pythonCmd = platform === 'win32' ? 'python' : 'python3';
    log.info(`Usando Python del sistema: ${pythonCmd}`);
    return { command: pythonCmd, args: [scriptPath] };
  }

  log.info('Modo producción: buscando ejecutable Python empaquetado');
  const resourcesPath = process.resourcesPath;
  const packagedDir = path.join(resourcesPath, 'python');
  const executableName = platform === 'win32' ? 'processor.exe' : 'processor';

  log.info(`Buscando en: ${packagedDir}`);
  log.info(`Nombre esperado: ${executableName}`);

  const executableCandidates = [
    path.join(packagedDir, executableName),
    path.join(packagedDir, 'bin', executableName),
    path.join(packagedDir, 'dist', executableName),
    path.join(packagedDir, 'dist', 'processor', executableName),
    path.join(packagedDir, 'processor', executableName)
  ];

  for (const candidate of executableCandidates) {
    log.info(`Verificando: ${candidate}`);
    if (fs.existsSync(candidate)) {
      log.info(`✅ Ejecutable encontrado: ${candidate}`);
      return { command: candidate, args: [] };
    }
    log.info(`❌ No encontrado: ${candidate}`);
  }

  const packagedScript = path.join(packagedDir, 'processor.py');
  log.warn(`Ejecutable no localizado. Intentando usar script Python: ${packagedScript}`);

  if (fs.existsSync(packagedScript)) {
    log.info(`✅ Script Python encontrado: ${packagedScript}`);
    const pythonCmd = platform === 'win32' ? 'python' : 'python3';
    return { command: pythonCmd, args: [packagedScript] };
  }

  log.error(`❌ Script Python no encontrado en ${packagedScript}`);
  throw new Error(`No se pudo encontrar un procesador Python empaquetado en ${packagedDir}`);
}

function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: 900,
    height: 600,
    minWidth: 820,
    minHeight: 520,
    show: false,
    title: 'Cartera San Benito',
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

      // Verificar que el comando existe antes de ejecutarlo
      if (!fs.existsSync(command) && !command.includes('python')) {
        const error = new Error(`El ejecutable no existe: ${command}`);
        log.error(error.message);
        reject(error);
        return;
      }

      const pyProcess = spawn(command, allArgs, { 
        stdio: ['ignore', 'pipe', 'pipe'],
        shell: process.platform === 'win32' // Usar shell en Windows para resolver comandos
      });

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
        log.error('Error al ejecutar Python:', error);
        let errorMessage = error.message;
        
        if (error.code === 'ENOENT') {
          errorMessage = `No se pudo encontrar el ejecutable: ${command}. Verifica que Python esté instalado correctamente.`;
        } else if (error.code === 9009) {
          errorMessage = `Error 9009: No se pudo ejecutar el comando Python. Verifica que Python esté en el PATH del sistema.`;
        }
        
        sendToRenderer('log-message', { level: 'error', message: errorMessage });
        reject(new Error(errorMessage));
      });

      pyProcess.on('close', code => {
        if (code === 0) {
          log.info(`Procesamiento finalizado. Archivo listo en ${outputPath}`);
          resolve(outputPath);
        } else {
          let errorMessage = `Python finalizó con código ${code}`;
          
          if (code === 9009) {
            errorMessage = `Error 9009: No se pudo encontrar o ejecutar Python. Verifica la instalación.`;
          } else if (code === 1) {
            errorMessage = `Error en el procesamiento Python. Revisa el archivo de entrada.`;
          }
          
          log.error(errorMessage);
          const err = new Error(errorMessage);
          err.code = code;
          reject(err);
        }
      });
    } catch (err) {
      log.error('Error en runPythonProcessor:', err);
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

  ipcMain.handle('updater:check-for-updates', async () => {
    try {
      const { autoUpdater } = require('electron-updater');
      log.info('Verificación manual de actualizaciones solicitada');
      const result = await autoUpdater.checkForUpdatesAndNotify();
      return { success: true, result };
    } catch (error) {
      log.error('Error en verificación manual de actualizaciones:', error);
      return { success: false, error: error.message };
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
  
  // Configurar opciones del autoUpdater
  updater.autoDownload = false; // No descargar automáticamente
  updater.autoInstallOnAppQuit = true; // Instalar al cerrar la app

  updater.on('update-available', (info) => {
    log.info('Actualización disponible:', info);
    sendToRenderer('update-event', { 
      type: 'available', 
      version: info.version,
      releaseNotes: info.releaseNotes 
    });
    
    // Preguntar al usuario si quiere descargar
    dialog.showMessageBox(mainWindow, {
      type: 'info',
      title: 'Actualización Disponible',
      message: `Se encontró una nueva versión (${info.version}). ¿Deseas descargarla ahora?`,
      buttons: ['Descargar', 'Más tarde'],
      defaultId: 0
    }).then((result) => {
      if (result.response === 0) {
        updater.downloadUpdate();
      }
    });
  });

  updater.on('update-not-available', (info) => {
    log.info('No hay actualizaciones disponibles:', info);
    sendToRenderer('update-event', { type: 'not-available' });
  });

  updater.on('error', error => {
    log.error('Error en autoUpdater:', error);
    sendToRenderer('update-event', { 
      type: 'error', 
      message: error?.message ?? String(error) 
    });
  });

  updater.on('download-progress', progressObj => {
    log.info('Progreso de descarga:', progressObj);
    sendToRenderer('update-event', { 
      type: 'progress', 
      progress: progressObj 
    });
  });

  updater.on('update-downloaded', (info) => {
    log.info('Actualización descargada, lista para instalar:', info);
    sendToRenderer('update-event', { 
      type: 'downloaded',
      version: info.version 
    });
    
    // Preguntar al usuario si quiere instalar ahora
    dialog.showMessageBox(mainWindow, {
      type: 'info',
      title: 'Actualización Descargada',
      message: 'La actualización se ha descargado. ¿Deseas reiniciar la aplicación ahora para instalarla?',
      buttons: ['Reiniciar ahora', 'Más tarde'],
      defaultId: 0
    }).then((result) => {
      if (result.response === 0) {
        updater.quitAndInstall();
      }
    });
  });
}

app.whenReady().then(() => {
  createMainWindow();
  registerIpcHandlers();

  log.info(`Aplicación inicializada. Versión ${app.getVersion()}`);

  try {
    const { autoUpdater } = require('electron-updater');
    setupAutoUpdater(autoUpdater);
    
    // Verificar actualizaciones al iniciar (solo en producción)
    if (app.isPackaged) {
      log.info('Verificando actualizaciones...');
      autoUpdater.checkForUpdatesAndNotify().catch(error => {
        log.error('Fallo al buscar actualizaciones:', error);
        sendToRenderer('update-event', { 
          type: 'error', 
          message: `Error al buscar actualizaciones: ${error.message}` 
        });
      });
    } else {
      log.info('AutoUpdater deshabilitado en modo desarrollo');
    }
  } catch (error) {
    log.warn('AutoUpdater deshabilitado o no disponible:', error);
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
