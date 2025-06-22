// [CODESENSEI] Versión refactorizada para procesamiento por lotes (batch processing)

import axios from 'axios';
import config from '../config/index.js';
import logger from '../utils/logger.js';

/**
 * Envía un lote de objetos de comentarios a N8N para su procesamiento.
 * @param {object} batchOfComments - Objeto donde cada clave es un nombre de carpeta y el valor es el objeto de comentarios.
 * @returns {Promise<object>} - Resuelve al objeto con todos los comentarios corregidos, o al objeto original si falla.
 */
async function correctAllCommentsInBatch(batchOfComments) {
  const webhookUrl = config.services.N8N_WEBHOOK_URL;

  if (!webhookUrl) {
    logger.warn("URL de webhook N8N no configurada. Se usarán los comentarios originales.");
    return batchOfComments;
  }

  if (!batchOfComments || Object.keys(batchOfComments).length === 0) {
    logger.info("No hay lotes de comentarios para enviar a N8N.");
    return batchOfComments;
  }

  logger.info(`Enviando un lote de ${Object.keys(batchOfComments).length} carpetas a N8N...`);

  try {
    const response = await axios.post(webhookUrl, batchOfComments, {
      headers: { 'Content-Type': 'application/json' },
      timeout: 120000, // [CODESENSEI] Aumentado el timeout a 120s para dar tiempo a procesar el lote.
    });
    
    if (response.data && typeof response.data === 'object') {
      logger.info("Lote de comentarios procesado y decodificado de N8N exitosamente.");
      return response.data;
    }
    logger.warn(`N8N respondió con ${response.status} OK pero sin datos válidos. Se usarán los comentarios originales.`);
  } catch (error) {
    // ... (la lógica de manejo de errores de axios se mantiene igual) ...
    if (axios.isAxiosError(error)) {
        logger.error(`Error de Axios al contactar N8N: ${error.message}`);
        // ...
    } else {
        logger.error(`Error inesperado durante la comunicación con N8N: ${error.message}`);
    }
  }

  logger.warn("Se usarán los comentarios originales para todo el lote debido a un fallo con N8N.");
  return batchOfComments;
}

export { correctAllCommentsInBatch };
