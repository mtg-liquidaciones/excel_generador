// src/utils/logger.js
import winston from 'winston';
import 'winston-daily-rotate-file';
import path from 'path';
import config from '../config/index.js';

const { combine, timestamp, printf, colorize, errors } = winston.format;
const { logging: loggingConfig } = config;

const consoleLogFormat = combine(
  colorize(),
  timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
  printf(({ level, message, timestamp: ts, ...metadata }) => {
    let msg = `${ts} ${level}: ${message}`;
    if (metadata && Object.keys(metadata).length > 0 && metadata.error instanceof Error) {
        msg += `\n${metadata.error.stack}`;
    }
    return msg;
  }),
  errors({ stack: true })
);

const fileLogFormat = combine(
  timestamp(),
  errors({ stack: true }),
  loggingConfig.formatType === 'json' 
    ? winston.format.json() 
    : printf(({ level, message, timestamp: ts, stack }) => `${ts} [${level.toUpperCase()}]: ${message}${stack ? `\n${stack}` : ''}`)
);

const transports = [];

if (loggingConfig.consoleEnabled) {
  transports.push(new winston.transports.Console({
    level: loggingConfig.level,
    format: consoleLogFormat,
  }));
}

if (loggingConfig.fileEnabled) {
  transports.push(new winston.transports.DailyRotateFile({
    level: loggingConfig.level,
    dirname: path.join(process.cwd(), loggingConfig.directory),
    filename: `${path.parse(loggingConfig.fileName).name}-%DATE%${path.parse(loggingConfig.fileName).ext}`,
    datePattern: 'YYYY-MM-DD',
    zippedArchive: true,
    maxSize: loggingConfig.maxSizeBytes,
    maxFiles: `${loggingConfig.maxFilesBackup}d`,
    format: fileLogFormat,
  }));
}

const logger = winston.createLogger({
  level: loggingConfig.level,
  transports,
  exitOnError: false,
});

export default logger;