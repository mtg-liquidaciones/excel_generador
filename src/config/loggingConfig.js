// src/config/loggingConfig.js

// --- CONFIGURACIÓN DE LOGGING ---

// Nivel de log. Para Winston, los niveles comunes son:
// error, warn, info, http, verbose, debug, silly.
// Se puede configurar mediante una variable de entorno.
const LOG_LEVEL = process.env.LOG_LEVEL || 'info';

// Directorio donde se guardarán los archivos de log, relativo a la raíz del proyecto.
const LOG_DIRECTORY = process.env.LOG_DIRECTORY || 'logs';

// Nombre del archivo de log principal.
const LOG_FILE_NAME = process.env.LOG_FILE_NAME || 'excel_generator_service.log';

// Formato del log. Las librerías como Winston tienen sus propios sistemas de formato.
// 'simple': un formato de texto legible.
// 'json': formato JSON estructurado, útil para sistemas de recolección de logs.
// La implementación del formato se hará en el módulo del logger.
const LOG_FORMAT_TYPE = process.env.LOG_FORMAT_TYPE || 'simple';

// Tamaño máximo del archivo de log antes de rotar (en bytes).
// Ejemplo: 10 * 1024 * 1024 para 10MB.
const LOG_MAX_SIZE_BYTES = parseInt(process.env.LOG_MAX_SIZE_BYTES, 10) || (10 * 1024 * 1024);

// Número máximo de archivos de log a mantener después de la rotación.
const LOG_MAX_FILES_BACKUP = parseInt(process.env.LOG_MAX_FILES_BACKUP, 10) || 5;

// Habilitar/deshabilitar el logging a la consola.
// Útil para desactivarlo en producción si solo se quiere loguear a archivo,
// o mantenerlo para ver logs en tiempo real.
const CONSOLE_LOGGING_ENABLED = process.env.CONSOLE_LOGGING_ENABLED ? (process.env.CONSOLE_LOGGING_ENABLED === 'true') : true;

// Habilitar/deshabilitar el logging a archivo.
const FILE_LOGGING_ENABLED = process.env.FILE_LOGGING_ENABLED ? (process.env.FILE_LOGGING_ENABLED === 'true') : true;


const loggingConfig = {
  level: LOG_LEVEL,
  directory: LOG_DIRECTORY,
  fileName: LOG_FILE_NAME,
  formatType: LOG_FORMAT_TYPE, // 'simple' o 'json'
  maxSizeBytes: LOG_MAX_SIZE_BYTES,
  maxFilesBackup: LOG_MAX_FILES_BACKUP,
  consoleEnabled: CONSOLE_LOGGING_ENABLED,
  fileEnabled: FILE_LOGGING_ENABLED,
};

export default loggingConfig;