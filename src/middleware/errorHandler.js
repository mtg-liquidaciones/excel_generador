// src/middleware/errorHandler.js

// import logger from '../utils/logger.js'; // Asumiremos que tendrás un logger configurado
// import config from '../config/index.js'; // Para acceder a process.env.NODE_ENV

// Placeholder logger - reemplaza con tu instancia de logger real
const logger = {
  error: (message, errorDetails) => {
    console.error(`[ERROR-HANDLER] ${message}`, errorDetails ? JSON.stringify(errorDetails, null, 2) : '');
    if (errorDetails && errorDetails.stack && (process.env.NODE_ENV === 'development' || !(errorDetails.isOperational))) {
      console.error(errorDetails.stack);
    }
  },
};

// El NODE_ENV se establece fuera de la aplicación (ej. al iniciar el script: NODE_ENV=production node app.js)
// o a través de un archivo .env cargado muy temprano.
const NODE_ENV = process.env.NODE_ENV || 'development';

/**
 * Middleware de manejo de errores de Express.
 * Este debe ser el ÚLTIMO middleware añadido a la aplicación.
 *
 * @param {Error} err - El objeto error.
 * @param {import('express').Request} req - El objeto Request de Express.
 * @param {import('express').Response} res - El objeto Response de Express.
 * @param {import('express').NextFunction} next - La función Next de Express (eslint lo marca si no se usa, pero es parte de la firma).
 */
// eslint-disable-next-line no-unused-vars
function errorHandler(err, req, res, next) {
  // Registrar el error con detalles en el servidor
  logger.error(`Error no manejado para ${req.method} ${req.originalUrl}: ${err.message}`, {
    name: err.name,
    statusCode: err.statusCode,
    isOperational: err.isOperational,
    // stack: err.stack, // El logger placeholder ya maneja el stack
    // Puedes añadir más propiedades del error si son relevantes
  });

  let statusCode = err.statusCode || 500;
  let clientMessage = "Ocurrió un error interno inesperado en el servidor.";

  // Manejo de tipos de error específicos
  if (err.name === "TimeoutError") { // Error personalizado de timeout (ej. en excelRoutes)
    statusCode = 504; // Gateway Timeout
    clientMessage = err.message || "La operación excedió el tiempo límite.";
  } else if (err.isAxiosError) { // Errores de Axios (comunicación con N8N)
    statusCode = err.response?.status || 502; // Bad Gateway o el status de la respuesta de N8N
    clientMessage = `Error al comunicarse con un servicio externo.`;
    // No exponer err.message de axios directamente al cliente en producción.
  } else if (err.name === 'SyntaxError' && err.status && err.status >= 400 && err.status < 500 && err.message.toLowerCase().includes('json')) {
    // Errores de parseo de JSON del body (de express.json())
    statusCode = 400; // Bad Request
    clientMessage = "La solicitud contiene JSON malformado.";
  } else if (err.isOperational) {
    // Para errores operacionales personalizados cuya información es seguro mostrar
    clientMessage = err.message;
  }


  // En producción, para errores 500 no operacionales, no enviar detalles al cliente.
  if (NODE_ENV === 'production' && statusCode === 500 && !err.isOperational) {
    clientMessage = "Ocurrió un error interno inesperado. Por favor, inténtelo de nuevo más tarde.";
  }

  // En desarrollo, podríamos querer enviar más información
  if (NODE_ENV === 'development') {
    // No sobrescribir mensajes específicos ya definidos
    if (statusCode === 500 && clientMessage.startsWith("Ocurrió un error interno")) {
        clientMessage = err.message || clientMessage;
    }
  }

  res.status(statusCode).json({
    status: "error",
    message: clientMessage,
    // Solo enviar el stack en desarrollo y si no es un error "controlado" muy genérico
    ...(NODE_ENV === 'development' && { errorName: err.name, stack: err.stack?.split('\n') })
  });
}

export default errorHandler;