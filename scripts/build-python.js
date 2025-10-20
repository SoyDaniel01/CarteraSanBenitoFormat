#!/usr/bin/env node
/**
 * Build the Python processor into a standalone executable using PyInstaller.
 */
const { spawnSync } = require('child_process');
const path = require('path');
const fs = require('fs');

const projectRoot = path.resolve(__dirname, '..');
const pythonDir = path.join(projectRoot, 'python');
const processorScript = path.join(pythonDir, 'processor.py');

function resolvePythonInterpreter() {
  const venvDir = path.join(projectRoot, '.venv');
  const candidates = process.platform === 'win32'
    ? [
        path.join(venvDir, 'Scripts', 'python.exe'),
        path.join(venvDir, 'Scripts', 'python'),
      ]
    : [
        path.join(venvDir, 'bin', 'python3'),
        path.join(venvDir, 'bin', 'python'),
      ];

  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }

  return process.platform === 'win32' ? 'python' : 'python3';
}

function ensurePyInstaller(pythonExecutable) {
  const check = spawnSync(
    pythonExecutable,
    ['-m', 'PyInstaller', '--version'],
    { cwd: projectRoot, stdio: 'ignore' }
  );

  if (check.status === 0) {
    return;
  }

  console.log('PyInstaller no encontrado. Instalando en el entorno actual...');
  const install = spawnSync(
    pythonExecutable,
    ['-m', 'pip', 'install', '--upgrade', 'pyinstaller'],
    { cwd: projectRoot, stdio: 'inherit' }
  );

  if (install.status !== 0) {
    throw new Error('No se pudo instalar PyInstaller.');
  }
}

function runPyInstaller(pythonExecutable) {
  const distPath = path.join(pythonDir, 'dist');
  const workPath = path.join(pythonDir, 'build');

  const args = [
    '-m',
    'PyInstaller',
    processorScript,
    '--onefile',
    '--name',
    'processor',
    '--distpath',
    distPath,
    '--workpath',
    workPath,
    '--clean'
  ];

  const result = spawnSync(pythonExecutable, args, { cwd: projectRoot, stdio: 'inherit' });

  if (result.status !== 0) {
    throw new Error('Falló la construcción del procesador de Python.');
  }

  const executableName = process.platform === 'win32' ? 'processor.exe' : 'processor';
  const outputPath = path.join(distPath, executableName);

  if (!fs.existsSync(outputPath)) {
    throw new Error(`No se encontró el ejecutable generado en ${outputPath}`);
  }

  console.log(`Procesador Python empaquetado correctamente: ${outputPath}`);

  return outputPath;
}

function main() {
  if (!fs.existsSync(processorScript)) {
    console.error('No se encontró python/processor.py. Aborting.');
    process.exit(1);
  }

  const pythonExecutable = resolvePythonInterpreter();
  ensurePyInstaller(pythonExecutable);
  runPyInstaller(pythonExecutable);
}

main();
