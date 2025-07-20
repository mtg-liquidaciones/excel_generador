// src/config/externalServicesConfig.js

// [CODESENSEI] Eliminado el valor de respaldo. La configuración debe venir del entorno (.env).
const N8N_WEBHOOK_URL = process.env.N8N_WEBHOOK_URL;

const externalServicesConfig = {
  N8N_WEBHOOK_URL,
};

export default externalServicesConfig;