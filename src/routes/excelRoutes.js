// src/routes/excelRoutes.js
// [CODESENSEI] Versión final compatible con la nueva estructura de JSON.

import express from 'express';
import fs from 'fs/promises';
import { generateFullExcelForPath } from '../services/excelProcessingService.js';
import config from '../config/index.js';
import logger from '../utils/logger.js';

const router = express.Router();

class TimeoutError extends Error {
  constructor(message) {
    super(message);
    this.name = "TimeoutError";
  }
}

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

router.post('/generar_excel', async (req, res, next) => {
  try {
    logger.info(`POST /generar_excel request received from ${req.ip}`);
    
    // [CODESENSEI] Cambio 1: Ahora trabajamos con todo el cuerpo de la petición.
    const requestBody = req.body;
    const { uri: projectPath } = requestBody; // Usamos la clave "uri" que te envía la desarrolladora.

    if (!projectPath) {
      logger.warn("'uri' not provided in JSON payload.");
      return res.status(400).json({ status: "error", message: "'uri' is required." });
    }

    const stats = await fs.stat(projectPath);
    if (!stats.isDirectory()) {
      logger.warn(`'uri' (${projectPath}) is not a valid directory.`);
      return res.status(404).json({ status: "error", message: `Project path is not a valid directory: ${projectPath}` });
    }

    logger.info(`Project path received: ${projectPath}`);
    const generationTimeoutMs = config.app.GENERATION_TIMEOUT_SECONDS * 1000;

    // [CODESENSEI] Cambio 2: Pasamos el 'requestBody' COMPLETO a la función.
    const excelGenerationPromise = generateFullExcelForPath(requestBody);
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
    }
  } catch (error) {
    // [CODESENSEI] Pasamos cualquier error a nuestro manejador centralizado.
    next(error);
  }
});

export default router;