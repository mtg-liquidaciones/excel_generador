// src/routes/excelRoutes.js

import express from 'express';
import fs from 'fs/promises';
// import path from 'path'; // path no se usa directamente aquí, pero es bueno tenerlo si se expande
import { generateFullExcelForPath } from '../services/excelProcessingService.js';
import config from '../config/index.js';
// import logger from '../utils/logger.js';

// Placeholder logger
const logger = {
  info: (message) => console.log(`[INFO] excelRoutes: ${message}`),
  warn: (message) => console.warn(`[WARN] excelRoutes: ${message}`),
  error: (message, error) => console.error(`[ERROR] excelRoutes: ${message}`, error || ''),
};

const router = express.Router();

/**
 * Custom TimeoutError class
 */
class TimeoutError extends Error {
  constructor(message) {
    super(message);
    this.name = "TimeoutError";
  }
}

/**
 * Wraps an async operation with a timeout.
 * @param {Promise<any>} promise - The promise to race against timeout.
 * @param {number} timeoutMs - Timeout in milliseconds.
 * @param {string} operationName - Name of the operation for logging.
 * @returns {Promise<any>} - The result of the original promise or throws TimeoutError.
 */
// 👇 CORRECCIÓN AQUÍ: Añadido espacio después de 'function'
function asyncOperationWithTimeout(promise, timeoutMs, operationName = "Async Operation") {
  let timeoutHandle;
  const timeoutPromise = new Promise((_, reject) => {
    timeoutHandle = setTimeout(() => {
      logger.error(`${operationName} timed out after ${timeoutMs / 1000} seconds.`);
      reject(new TimeoutError(`${operationName} exceeded timeout of ${timeoutMs / 1000} seconds.`));
    }, timeoutMs);
  });

  return Promise.race([promise, timeoutPromise])
    .finally(() => {
      clearTimeout(timeoutHandle);
    });
}

// POST /api/excel/generar_excel
router.post('/generar_excel', async (req, res) => {
  logger.info(`POST /generar_excel request received from ${req.ip}`);

  if (!req.body || typeof req.body !== 'object') {
    logger.warn("Request body is not JSON or is missing.");
    return res.status(400).json({ status: "error", message: "Request body must be JSON." });
  }

  const { ruta_proyecto: projectPath } = req.body;

  if (!projectPath) {
    logger.warn("'ruta_proyecto' not provided in JSON payload.");
    return res.status(400).json({ status: "error", message: "'ruta_proyecto' is required." });
  }

  try {
    const stats = await fs.stat(projectPath);
    if (!stats.isDirectory()) {
      logger.warn(`'ruta_proyecto' (${projectPath}) is not a valid directory.`);
      return res.status(404).json({ status: "error", message: `Project path is not a valid directory: ${projectPath}` });
    }
  } catch (error) {
    logger.warn(`Error accessing 'ruta_proyecto' (${projectPath}): ${error.message}`);
    return res.status(404).json({ status: "error", message: `Project path does not exist or is not accessible: ${projectPath}` });
  }

  logger.info(`Project path received: ${projectPath}`);

  const generationTimeoutMs = config.app.GENERATION_TIMEOUT_SECONDS * 1000;

  try {
    const excelGenerationPromise = generateFullExcelForPath(projectPath);
    const result = await asyncOperationWithTimeout(
      excelGenerationPromise,
      generationTimeoutMs,
      `Excel generation for '${projectPath}'`
    );

    if (result.success) {
      logger.info(`Excel generation successful for '${projectPath}': ${result.message}. File: ${result.filePath}`);
      return res.status(200).json({
        status: "success",
        message: result.message,
        file_path: result.filePath,
      });
    } else {
      logger.error(`Excel generation failed or was partial for '${projectPath}': ${result.message}`);
      return res.status(500).json({ status: "error", message: result.message || "Excel generation failed." });
    }
  } catch (error) {
    if (error instanceof TimeoutError) {
      logger.error(`Timeout (${generationTimeoutMs / 1000}s) reached during Excel generation for '${projectPath}'.`);
      return res.status(504).json({
        status: "error",
        message: `Timeout: Excel generation exceeded ${generationTimeoutMs / 1000} seconds.`,
      });
    } else {
      logger.error(`Unexpected error during Excel generation for '${projectPath}': ${error.message}`, error);
      return res.status(500).json({
        status: "error",
        message: "An internal server error occurred during Excel generation.",
      });
    }
  }
});

export default router;