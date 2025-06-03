// src/services/n8nService.js

import axios from 'axios';
import config from '../config/index.js'; // Importa la configuración principal
import logger from '../utils/logger.js'; // <--- Importamos el logger real aquí

/**
 * Envía los comentarios a un webhook de N8N para su procesamiento/corrección.
 * Si la URL del webhook no está configurada o si hay un error en la comunicación,
 * se devolverán los comentarios originales.
 * @param {object} originalComments - El objeto con los comentarios originales.
 * @returns {Promise<object>} - Una promesa que resuelve al objeto de comentarios (corregidos o originales).
 */
async function correctCommentsViaN8N(originalComments) {
  const webhookUrl = config.services.N8N_WEBHOOK_URL;

  if (!webhookUrl) {
    logger.warn("URL de webhook N8N no configurada. Se usarán los comentarios originales.");
    return originalComments;
  }

  if (!originalComments || Object.keys(originalComments).length === 0) {
    logger.info("No se proporcionaron comentarios originales para enviar a N8N. Se devolverán tal cual.");
    return originalComments || {}; // Devuelve un objeto vacío si originalComments es null/undefined
  }

  logger.info(`Enviando ${Object.keys(originalComments).length} comentarios a N8N (${webhookUrl})...`);

  try {
    const response = await axios.post(
      webhookUrl,
      originalComments,
      {
        headers: { 'Content-Type': 'application/json' },
        timeout: 60000, // 60 segundos de timeout (axios usa milisegundos)
      }
    );

    // Axios por defecto considera errores los status fuera del rango 2xx,
    // por lo que si llegamos aquí, el status es 2xx.
    logger.info(`Respuesta de N8N - Código de estado: ${response.status}`);

    // Verificar si la respuesta contiene datos y si es un objeto
    if (response.data && typeof response.data === 'object') {
      logger.info("Comentarios decodificados de N8N exitosamente.");
      return response.data;
    } else if (response.status >= 200 && response.status < 300) {
      // El status es 2xx pero response.data no es un objeto o está vacío
      logger.warn(`N8N respondió con ${response.status} OK pero el cuerpo de la respuesta estaba vacío o no era un objeto JSON válido. Usando comentarios originales.`);
      logger.debug(`Respuesta cruda de N8N (datos): ${JSON.stringify(response.data)}`);
    } else {
      // Este caso es menos probable con Axios si ya manejó el error de status, pero por si acaso.
      logger.warn(`Respuesta inesperada de N8N (status ${response.status}). Usando comentarios originales.`);
    }

  } catch (error) {
    if (axios.isAxiosError(error)) {
      logger.error(`Error de Axios al contactar N8N: ${error.message}`);
      if (error.response) {
        // La solicitud se realizó y el servidor respondió con un código de estado
        // que cae fuera del rango de 2xx
        logger.error(`Respuesta de N8N - Status: ${error.response.status}`);
        logger.error(`Respuesta de N8N - Datos: ${JSON.stringify(error.response.data)}`);
      } else if (error.request) {
        // La solicitud se realizó pero no se recibió respuesta (ej. timeout)
        logger.error("La solicitud a N8N fue realizada pero no se recibió respuesta.");
      } else {
        // Algo sucedió al configurar la solicitud que provocó un error
        logger.error(`Error al configurar la solicitud a N8N: ${error.message}`);
      }
    } else {
      // Otro tipo de error
      logger.error(`Error inesperado durante la comunicación con N8N: ${error.message}`);
    }
  }

  logger.warn("Se usarán los comentarios originales debido a un problema con N8N o una respuesta inesperada/vacía.");
  return originalComments;
}

export { correctCommentsViaN8N };