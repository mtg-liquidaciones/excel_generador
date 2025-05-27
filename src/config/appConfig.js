// src/config/appConfig.js

// Es una buena práctica inicializar dotenv en el punto de entrada principal de la aplicación (ej. app.js)
// para que process.env esté poblado antes de que se carguen estos módulos de configuración.
// Ejemplo: import dotenv from 'dotenv'; dotenv.config();

// --- CONFIGURACIÓN DEL SERVICIO ---

// Puerto en el que escuchará el servicio.
// Intenta leerlo desde las variables de entorno, si no, usa un valor por defecto.
const SERVICE_PORT = parseInt(process.env.SERVICE_PORT, 10) || 9898;

// Tiempo máximo en segundos que el servicio esperará para la generación de un archivo Excel.
// Intenta leerlo desde las variables de entorno, si no, usa un valor por defecto.
const GENERATION_TIMEOUT_SECONDS = parseInt(process.env.GENERATION_TIMEOUT_SECONDS, 10) || 180;

// Podríamos añadir otras configuraciones específicas de la aplicación aquí, si las hubiera.
// Por ejemplo, si el nombre del servicio o la versión se gestionan desde aquí:
// const SERVICE_NAME = process.env.SERVICE_NAME || 'ExcelGeneratorService';
// const SERVICE_VERSION = process.env.SERVICE_VERSION || '1.0.0';

const appConfig = {
  SERVICE_PORT,
  GENERATION_TIMEOUT_SECONDS,
  // SERVICE_NAME, // Descomentar si se añade
  // SERVICE_VERSION, // Descomentar si se añade
};

export default appConfig;