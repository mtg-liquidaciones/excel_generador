// src/app.js

// 1. Cargar variables de entorno (lo antes posible)
import dotenv from 'dotenv'; // <--- ASEGURADO QUE ESTÁ SOLO UNA VEZ
dotenv.config(); // Carga variables desde el archivo .env

// 2. Importar dependencias principales
import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url'; // Para obtener __dirname en ES modules
import config from './config/index.js';
import excelRoutes from './routes/excelRoutes.js';
import errorHandler from './middleware/errorHandler.js';
import { ensureDirectoryExists } from './utils/fileSystemUtils.js';
// import logger from './utils/logger.js'; // Asumimos que crearás este módulo

// Placeholder logger - Reemplázalo con la importación de tu logger configurado
const logger = {
  info: (message) => console.log(`[INFO] App: ${message}`),
  error: (message, error) => console.error(`[ERROR] App: ${message}`, error || ''),
  debug: (message) => console.log(`[DEBUG] App: ${message}`),
};

// Obtener __dirname en ES modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// 3. Inicializar la aplicación Express
const app = express();

// 4. Asegurar que el directorio de logs exista
// (Deberás crear src/utils/logger.js que use loggingConfig para la ruta completa)
ensureDirectoryExists(path.join(process.cwd(), config.logging.directory))
  .then(() => logger.info(`Directorio de logs asegurado: ${config.logging.directory}`))
  .catch(error => logger.error(`No se pudo asegurar el directorio de logs: ${error.message}`, error));

// 5. Configurar Middlewares globales
app.use(express.json({ limit: '1mb' })); // Para parsear cuerpos de solicitud JSON (ajusta el límite si es necesario)
app.use(express.urlencoded({ extended: true, limit: '1mb' })); // Para parsear cuerpos de solicitud URL-encoded

// Middleware para logging de peticiones (opcional, puedes usar morgan o tu logger)
app.use((req, res, next) => {
  logger.info(`${req.method} ${req.originalUrl} - IP: ${req.ip}`);
  res.on('finish', () => {
    logger.info(`${res.statusCode} ${req.method} ${req.originalUrl} - IP: ${req.ip}`);
  });
  next();
});

// Middleware para CORS (Cross-Origin Resource Sharing) - Descomentar si es necesario
// import cors from 'cors';
// app.use(cors()); // Configuración básica, o puedes pasar opciones: app.use(cors({ origin: 'https://tu-frontend.com' }))

// 6. Montar Rutas
// Todas las rutas definidas en excelRoutes estarán prefijadas con /api/excel
app.use('/api/excel', excelRoutes);

// Ruta raíz de prueba
app.get('/', (req, res) => {
  res.status(200).json({
    message: 'Servicio Generador de Excel funcionando correctamente.',
    status: 'OK',
    timestamp: new Date().toISOString(),
  });
});

// 7. Manejador para rutas no encontradas (404)
// Debe ir después de todas tus rutas definidas
app.use((req, res, next) => {
  const error = new Error(`Ruta no encontrada - ${req.originalUrl}`);
  // @ts-ignore // Para añadir statusCode a un Error estándar
  error.statusCode = 404;
  next(error); // Pasa el error al manejador de errores centralizado
});

// 8. Manejador de Errores Centralizado
// ¡Importante! Este debe ser el ÚLTIMO middleware que se añade.
app.use(errorHandler);

// 9. Iniciar el Servidor
const PORT = config.app.SERVICE_PORT;

app.listen(PORT, () => {
  logger.info(`Servicio Generador de Excel iniciado en el puerto ${PORT}`);
  logger.info(`Entorno actual: ${process.env.NODE_ENV || 'development'}`);
  // Puedes añadir más información al inicio si es útil
});

// Manejo de errores no capturados y promesas rechazadas (opcional pero recomendado)
process.on('unhandledRejection', (reason, promise) => {
  logger.error('Unhandled Rejection at:', { promise, reason });
  // Considera cerrar la aplicación de forma controlada en algunos casos:
  // process.exit(1);
});

process.on('uncaughtException', (error) => {
  logger.error('Uncaught Exception:', error);
  // Es crítico cerrar la aplicación aquí, ya que el estado puede ser corrupto
  process.exit(1);
});

export default app; // Útil para pruebas o si se usa con serverless/otros frameworks