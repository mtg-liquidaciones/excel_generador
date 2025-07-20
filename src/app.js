// src/app.js

// [CODESENSEI] Ya no cargamos dotenv aquí. Se carga centralmente en el módulo de config.
import express from 'express';
import path from 'path';
import config from './config/index.js'; // Este import ahora dispara la carga del .env
import excelRoutes from './routes/excelRoutes.js';
import errorHandler from './middleware/errorHandler.js';
import { ensureDirectoryExists } from './utils/fileSystemUtils.js';
import logger from './utils/logger.js';

const app = express();

ensureDirectoryExists(path.join(process.cwd(), config.logging.directory))
  .then(() => logger.info(`Directorio de logs asegurado: ${config.logging.directory}`))
  .catch(error => logger.error(`No se pudo asegurar el directorio de logs: ${error.message}`, error));

app.use(express.json({ limit: '1mb' }));
app.use(express.urlencoded({ extended: true, limit: '1mb' }));

app.use((req, res, next) => {
  logger.info(`${req.method} ${req.originalUrl} - IP: ${req.ip}`);
  next();
});

app.use('/api/excel', excelRoutes);

app.get('/', (req, res) => {
  res.status(200).json({
    message: 'Servicio Generador de Excel funcionando correctamente.',
    status: 'OK',
    timestamp: new Date().toISOString(),
  });
});

app.use((req, res, next) => {
  const error = new Error(`Ruta no encontrada - ${req.originalUrl}`);
  error.statusCode = 404;
  next(error);
});

app.use(errorHandler);

const PORT = config.app.SERVICE_PORT;

app.listen(PORT, () => {
  logger.info(`Servicio Generador de Excel iniciado en el puerto ${PORT}`);
  logger.info(`Entorno actual: ${process.env.NODE_ENV || 'development'}`);
});

process.on('unhandledRejection', (reason, promise) => {
  logger.error('Unhandled Rejection at:', { promise, reason });
});

process.on('uncaughtException', (error) => {
  logger.error('Uncaught Exception:', error);
  process.exit(1);
});

export default app;