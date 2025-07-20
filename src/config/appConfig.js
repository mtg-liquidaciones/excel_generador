// src/config/appConfig.js

const SERVICE_PORT = parseInt(process.env.SERVICE_PORT, 10) || 9898;
const GENERATION_TIMEOUT_SECONDS = parseInt(process.env.GENERATION_TIMEOUT_SECONDS, 10) || 180;

// [CODESENSEI] Añadidas nuevas variables para que sean configurables desde el .env.
const BODY_LIMIT = process.env.BODY_LIMIT || '10mb';
const CORS_ORIGIN = process.env.CORS_ORIGIN || '*';
const NODE_ENV = process.env.NODE_ENV || 'development';

const appConfig = {
  SERVICE_PORT,
  GENERATION_TIMEOUT_SECONDS,
  BODY_LIMIT,
  CORS_ORIGIN,
  NODE_ENV,
};

export default appConfig;