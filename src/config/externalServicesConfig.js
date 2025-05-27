// src/config/externalServicesConfig.js

// Recordatorio: dotenv debe estar inicializado en el punto de entrada principal
// de la aplicación (ej. app.js) para que process.env esté disponible.

// --- CONFIGURACIÓN PARA N8N ---

// URL del Webhook de N8N para la corrección de comentarios.
// Intenta leerla desde las variables de entorno; si no, usa el valor por defecto
// que tenías en tu config.py.
// Para producción, es altamente recomendable configurar esto SÓLO vía variables de entorno.
const N8N_WEBHOOK_URL =
  process.env.N8N_WEBHOOK_URL || "https://n8n.everytel.pe/webhook/ab3abc73-a8c1-4d2c-bd41-32dc373e80f2";

// Si hubiera otros servicios externos, sus configuraciones irían aquí.
// Ejemplo:
// const ANOTHER_SERVICE_API_KEY = process.env.ANOTHER_SERVICE_API_KEY || 'default_api_key';
// const ANOTHER_SERVICE_URL = process.env.ANOTHER_SERVICE_URL || 'https://api.anotherservice.com/v1';

const externalServicesConfig = {
  N8N_WEBHOOK_URL,
  // ANOTHER_SERVICE_API_KEY, // Descomentar si se añade
  // ANOTHER_SERVICE_URL,     // Descomentar si se añade
};

export default externalServicesConfig;