// src/config/index.js

// [CODESENSEI] SOLUCIÓN DEFINITIVA: Cargar .env aquí, ANTES que cualquier otra cosa.
import path from 'path';
import dotenv from 'dotenv';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
// Sube dos niveles desde /src/config para llegar a la raíz del proyecto
const projectRoot = path.resolve(__dirname, '..', '..'); 
const envPath = path.join(projectRoot, '.env');
dotenv.config({ path: envPath });

// Ahora que .env está cargado, importamos el resto de las configuraciones
import appConfig from './appConfig.js';
import loggingConfig from './loggingConfig.js';
import styleAndLayoutConfig from './styleAndLayoutConfig.js';
import externalServicesConfig from './externalServicesConfig.js';

// Combinar todas las configuraciones en un solo objeto
const config = {
  app: appConfig,
  logging: loggingConfig,
  excel: styleAndLayoutConfig,
  services: externalServicesConfig,
};

export default config;