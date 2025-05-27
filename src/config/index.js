// src/config/index.js

// Importar las configuraciones específicas
// Asumiremos que cada uno de estos archivos exportará un objeto por defecto
// o múltiples exportaciones nombradas según sea necesario.
import appConfig from './appConfig.js';
import loggingConfig from './loggingConfig.js';
import styleAndLayoutConfig from './styleAndLayoutConfig.js';
import externalServicesConfig from './externalServicesConfig.js';

// Combinar todas las configuraciones en un solo objeto para una fácil importación
// en otros módulos.
const config = {
  // Configuraciones de la aplicación (puerto, timeouts)
  app: appConfig,

  // Configuraciones de logging (Winston/Pino)
  logging: loggingConfig,

  // Configuraciones de estilos, diseño y mapeos para Excel,
  // y nombres de archivo/rutas base relevantes para la generación.
  excel: styleAndLayoutConfig,

  // Configuraciones para servicios externos (ej. N8N)
  services: externalServicesConfig,
};

// También podríamos optar por exportar individualmente si se prefiere
// acceder a una sección de configuración directamente:
export {
  appConfig,
  loggingConfig,
  styleAndLayoutConfig,
  externalServicesConfig
};

// Exportar el objeto de configuración combinado por defecto
export default config;