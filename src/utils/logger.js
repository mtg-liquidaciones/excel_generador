// src/utils/logger.js
import winston from 'winston';
import 'winston-daily-rotate-file'; // Para rotación de archivos de log
import config from '../config/index.js';

const logDir = config.logging.directory; // logs/
const logFileName = config.logging.fileName; // excel_generator_service.log
const logLevel = config.logging.level; // info
const logFormatType = config.logging.formatType; // simple o json
const logMaxSizeBytes = config.logging.maxSizeBytes; // 10MB
const logMaxFilesBackup = config.logging.maxFilesBackup; // 5
const consoleLoggingEnabled = config.logging.consoleEnabled;
const fileLoggingEnabled = config.logging.fileEnabled;

// Asegúrate de que el directorio de logs exista antes de inicializar Winston
// Esto ya se hace en app.js con ensureDirectoryExists, pero es bueno tenerlo en cuenta.
// La inicialización de Winston se hace en caliente, por lo que el directorio debe existir antes.

const transports = [];

// Transporte para consola
if (consoleLoggingEnabled) {
  transports.push(
    new winston.transports.Console({
      level: logLevel,
      format: winston.format.combine(
        winston.format.colorize(),
        winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
        winston.format.printf(info => `${info.timestamp} ${info.level}: ${info.message}${info.stack ? `\n${info.stack}` : ''}`)
      ),
    })
  );
}

// Transporte para archivo (con rotación)
if (fileLoggingEnabled) {
  transports.push(
    new winston.transports.DailyRotateFile({
      level: logLevel,
      dirname: logDir,
      filename: logFileName.replace('.log', '-%DATE%.log'), // e.g., excel_generator_service-2023-01-01.log
      datePattern: 'YYYY-MM-DD',
      zippedArchive: true, // Comprime los archivos de log antiguos
      maxSize: logMaxSizeBytes, // Tamaño máximo del archivo (ej. '20m' para 20MB)
      maxFiles: logMaxFilesBackup, // Número máximo de archivos a mantener (ej. '14d' para 14 días)
      format: logFormatType === 'json' ?
        winston.format.combine(
          winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
          winston.format.json()
        ) :
        winston.format.combine(
          winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
          winston.format.printf(info => `${info.timestamp} ${info.level}: ${info.message}${info.stack ? `\n${info.stack}` : ''}`)
        ),
    })
  );
}

const logger = winston.createLogger({
  levels: winston.config.npm.levels, // Usa los niveles estándar de npm (error, warn, info, etc.)
  format: winston.format.json(), // Formato por defecto para que los transportes lo sobrescriban si es necesario
  transports: transports,
  exitOnError: false, // No salga en caso de error en un transporte
});

// Para capturar excepciones no controladas y promesas rechazadas
// Esto es un fallback si no se capturan en app.js, pero app.js ya lo hace
// logger.exceptions.handle(
//   new winston.transports.File({ filename: path.join(logDir, 'exceptions.log') })
// );

// process.on('unhandledRejection', (reason, promise) => {
//   logger.error('Unhandled Rejection at:', promise, 'reason:', reason);
// });

export default logger;