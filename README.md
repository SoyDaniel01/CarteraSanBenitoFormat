# Cartera San Benito Format

Aplicación de escritorio multiplataforma basada en Electron que delega el procesamiento de archivos Excel a Python y soporta actualizaciones automáticas a través de GitHub Releases.

> **Nota:** El script Python procesa la hoja `Cartera_CxC_Det_Comple` (limpieza con pandas, totales y estilos con openpyxl) y deja preparadas las demás hojas para sus futuros flujos.

## Características principales

- UI amigable con drag & drop y botón de selección de archivos `.xlsx`.
- Comunicación segura entre Renderer y Main mediante preload con `contextBridge`.
- Ejecución de scripts Python dentro de un proceso hijo controlado y logueado.
- Logs centralizados vía `electron-log`, visibles en la aplicación y almacenados en disco.
- Configuración preparada para `electron-builder` + `electron-updater` con publicador GitHub.
- Estructura compatible con Windows y macOS (auto-paqueteo `.dmg` / `.exe`).

## Estructura del proyecto

```
├── src/
│   ├── main/
│   │   ├── main.js        # Proceso principal de Electron
│   │   └── preload.js     # API segura expuesta al Renderer
│   └── renderer/
│       ├── index.html     # Interfaz de usuario
│       ├── renderer.js    # Lógica del Renderer
│       └── styles.css     # Estilos base
├── python/
│   ├── processor.py       # Script de procesamiento (Cartera_CxC_Det_Comple)
│   └── requirements.txt   # Dependencias de Python
├── build/                 # Recursos para empaquetado (icons, etc.)
├── package.json
└── README.md
```

## Requisitos previos

- Node.js ≥ 18
- npm ≥ 9
- Python 3.x con `pip` (para desarrollo o reempaquetado del script)

## Configuración rápida

1. Instala dependencias JavaScript:
   ```bash
   npm install
   ```
2. Instala las dependencias de Python en tu entorno preferido:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate  # .venv\Scripts\activate en Windows
   pip install -r python/requirements.txt
   ```

## Desarrollo

Inicia la aplicación en modo desarrollo:
```bash
npm run dev
```

Esto lanza Electron apuntando a `src/main/main.js`. La UI permite arrastrar un archivo `.xlsx` o seleccionarlo mediante diálogo nativo. El procesamiento genera un archivo `<nombre>_processed.xlsx` en la misma carpeta.

## Empaquetado

`electron-builder` ya está configurado en `package.json`. Para generar artefactos:

- macOS:
  ```bash
  npm run dist:mac
  ```
- Windows:
  ```bash
  npm run dist:win
  ```

Los builds se generan en `dist/`. Personaliza iconos en `build/` según las guías de cada plataforma.

### Empaquetado de Python

El script `python/processor.py` puede convertirse en un ejecutable autónomo (incluye pandas y dependencias) con:

```bash
npm run build:python
```

El comando:

- Usa el intérprete de `.venv` si existe, o el Python global.
- Instala PyInstaller en caso de que falte.
- Genera un binario en `python/dist/processor[.exe]`.

Al ejecutar `npm run dist`, `electron-builder` copiará ese binario al paquete final y la app intentará usarlo automáticamente. Si no encuentra el ejecutable, volverá a ejecutar el script `.py` con el Python del sistema.

## Auto-actualizaciones

`electron-updater` está configurado para publicar en `SoyDaniel01/CarteraSanBenitoFormat`. Para activar las actualizaciones:

1. Crea un token de GitHub (si el repo es privado) y configúralo como `GH_TOKEN` en tu entorno.
2. Genera un release con `npm run dist`.
3. Sube los artefactos generados (`.exe`, `.dmg`, `.yml`) al release correspondiente.
4. Al iniciar la app en producción, `autoUpdater.checkForUpdatesAndNotify()` buscará y descargará nuevas versiones.

## Logging

- Logs del proceso principal: `electron-log` los escribe en el directorio por defecto del SO (`~/Library/Logs`, `%APPDATA%`, etc.).
- Logs de Python: todo `stdout/stderr` es capturado y mostrado en la UI, además de guardarse en los logs de Electron.

## Próximos pasos sugeridos

1. Completar los flujos restantes para las hojas auxiliares (Res_Comp, DSE, RSE, etc.).
2. Ejecutar `npm run build:python` en cada plataforma objetivo para producir el ejecutable definitivo.
3. Añadir pruebas automatizadas (por ejemplo, scripts Node para validar el pipeline).
4. Integrar firma de código si la distribución será externa (especialmente macOS).
