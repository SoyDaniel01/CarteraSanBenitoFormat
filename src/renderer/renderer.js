const dropzone = document.getElementById('dropzone');
const browseButton = document.getElementById('browseButton');
const convertButton = document.getElementById('convertButton');
const openOutputButton = document.getElementById('openOutputButton');
const fileInfo = document.getElementById('fileInfo');
const statusMessage = document.getElementById('statusMessage');

let selectedFilePath = null;
let processedOutputPath = null;
let isProcessing = false;

function setStatus(state, message) {
  statusMessage.classList.remove('ready', 'running', 'error', 'success');
  statusMessage.classList.add(state);
  statusMessage.textContent = message;
}

function setSelectedFile(filePath) {
  selectedFilePath = filePath;
  if (filePath) {
    fileInfo.textContent = filePath;
    convertButton.disabled = false;
    setStatus('ready', 'Archivo listo para convertir');
    console.info('[Renderer] Archivo seleccionado:', filePath);
  } else {
    fileInfo.textContent = 'Ningún archivo seleccionado';
    convertButton.disabled = true;
    openOutputButton.disabled = true;
    processedOutputPath = null;
    setStatus('ready', 'Listo para procesar');
  }
}

function sanitizeFile(file) {
  if (!file) return null;
  if (!file.name.toLowerCase().endsWith('.xlsx')) {
    setStatus('error', 'Solo se permiten archivos .xlsx');
    console.warn('[Renderer] Intento de carga con extensión inválida.');
    return null;
  }
  if (!file.path) {
    setStatus('error', 'No se pudo obtener la ruta del archivo.');
    return null;
  }
  return file.path;
}

function enableProcessingUi(running) {
  isProcessing = running;
  convertButton.disabled = running || !selectedFilePath;
  browseButton.disabled = running;
  dropzone.classList.toggle('disabled', running);
}

async function handleBrowse() {
  try {
    const filePath = await window.carteraApi.selectFile();
    if (filePath) {
      setSelectedFile(filePath);
    }
  } catch (error) {
    console.error('[Renderer] No se pudo seleccionar archivo:', error);
  }
}

async function handleConvert() {
  if (!selectedFilePath || isProcessing) {
    return;
  }

  enableProcessingUi(true);
  setStatus('running', 'Procesando archivo…');
  console.info('[Renderer] Iniciando procesamiento:', selectedFilePath);

  const result = await window.carteraApi.processFile(selectedFilePath);

  enableProcessingUi(false);

  if (result?.success) {
    processedOutputPath = result.outputPath;
    openOutputButton.disabled = !processedOutputPath;
    setStatus('success', 'Procesamiento completado correctamente');
    console.info('[Renderer] Archivo procesado:', result.outputPath);
  } else {
    processedOutputPath = null;
    openOutputButton.disabled = true;
    const errorMessage = result?.error ?? 'Error desconocido';
    setStatus('error', `Fallo al procesar: ${errorMessage}`);
    console.error('[Renderer] Fallo en el procesamiento:', errorMessage);
  }
}

async function handleOpenOutput() {
  if (!processedOutputPath) {
    return;
  }

  const success = await window.carteraApi.revealInFolder(processedOutputPath);
  if (!success) {
    console.error('[Renderer] No se pudo abrir la ubicación del archivo procesado.');
  }
}

function setupDragAndDrop() {
  dropzone.addEventListener('dragover', (event) => {
    event.preventDefault();
    dropzone.classList.add('dragover');
  });

  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('dragover');
  });

  dropzone.addEventListener('drop', (event) => {
    event.preventDefault();
    dropzone.classList.remove('dragover');

    const [file] = event.dataTransfer?.files ?? [];
    const filePath = sanitizeFile(file);

    if (filePath) {
      setSelectedFile(filePath);
    }
  });
}

function setupIpcSubscriptions() {
  window.carteraApi.onLogMessage((logEntry) => {
    if (logEntry?.message) {
      const level = logEntry.level ?? 'info';
      const prefix = `[Main][${level}]`;
      if (level === 'error') {
        console.error(prefix, logEntry.message);
      } else {
        console.info(prefix, logEntry.message);
      }
    }
  });

  window.carteraApi.onProcessingState((state) => {
    if (!state) return;

    switch (state.status) {
      case 'running':
        setStatus('running', `Procesando ${state.input}…`);
        break;
      case 'completed':
        processedOutputPath = state.output;
        openOutputButton.disabled = !processedOutputPath;
        setStatus('success', 'Archivo procesado con éxito');
        break;
      case 'failed':
        processedOutputPath = null;
        openOutputButton.disabled = true;
        setStatus('error', state.error ?? 'Fallo en el procesamiento');
        break;
      default:
        break;
    }
  });

  window.carteraApi.onUpdateEvent((update) => {
    if (!update) return;
    console.info('[Updater]', update);
  });
}

function init() {
  setupDragAndDrop();
  setupIpcSubscriptions();

  browseButton.addEventListener('click', handleBrowse);
  convertButton.addEventListener('click', handleConvert);
  openOutputButton.addEventListener('click', handleOpenOutput);

  console.info('[Renderer] Interfaz inicializada.');
}

document.addEventListener('DOMContentLoaded', init);
