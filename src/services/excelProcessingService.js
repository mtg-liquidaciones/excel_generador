// src/services/excelProcessingService.js

import ExcelJS from 'exceljs';
import path from 'path';
import fs from 'fs/promises';
import config from '../config/index.js';
import { createSheetActaResumen } from '../excelGeneration/sheetActaResumen.js';
import { createSheetMedicionesOTDR } from '../excelGeneration/sheetMedicionesOTDR.js';
import { createConformitySheetInstance } from '../excelGeneration/sheetConformidad.js';
import { correctCommentsViaN8N } from './n8nService.js';
import logger from '../utils/logger.js'; // <--- Importamos el logger real aquí


function sanitizeFileName(fileName) {
  if (!fileName) return 'default_filename';
  return fileName.replace(/[^a-z0-9_.-]+/gi, '_');
}

/**
 * Generates the complete Excel workbook for a given project path.
 * @param {string} mainProjectFolder - Absolute path to the main project folder.
 * @returns {Promise<{success: boolean, message: string, filePath: string|null}>}
 */
async function generateFullExcelForPath(mainProjectFolder) {
  logger.info(`Starting Excel generation for project in: ${mainProjectFolder}`);

  const mainDataJsonPath = path.join(mainProjectFolder, config.excel.NOMBRE_ARCHIVO_DATOS_PRINCIPAL);
  let generalData = {};
  try {
    const fileContent = await fs.readFile(mainDataJsonPath, 'utf-8');
    generalData = JSON.parse(fileContent);
    logger.info(`Main datos.json loaded from ${mainDataJsonPath}.`);
  } catch (error) {
    logger.error(`Error loading or parsing ${mainDataJsonPath}: ${error.message}`, error);
  }

  const workbook = new ExcelJS.Workbook();

  logger.info("Creating 'Acta Resumen Pext' sheet...");
  const resumenSheet = workbook.addWorksheet("Acta Resumen Pext", {
    views: [{ showGridLines: false }]
  });
  await createSheetActaResumen(resumenSheet, generalData, config.excel.RUTA_LOGO_GTD);
  logger.info("'Acta Resumen Pext' created.");

  let validSubfoldersProcessedCount = 0;

  try {
    const projectContents = await fs.readdir(mainProjectFolder, { withFileTypes: true });

    for (const dirent of projectContents) {
      if (!dirent.isDirectory()) continue;

      const subfolderName = dirent.name;
      const subfolderPath = path.join(mainProjectFolder, subfolderName);
      const conformitySheetTitleTemplate = config.excel.FOLDER_TO_TITLE_MAP[subfolderName];

      if (conformitySheetTitleTemplate) {
        validSubfoldersProcessedCount++;
        logger.info(`Processing subfolder for conformity sheet: ${subfolderName}`);

        const sheetName = subfolderName.substring(0, 31);
        const conformitySheet = workbook.addWorksheet(sheetName, {
          views: [{ showGridLines: false }]
        });

        const colWidths = config.excel.CONFORMIDAD_SHEET_CONFIG.anchos_columnas_char;
        if (colWidths) {
          Object.entries(colWidths).forEach(([colLetter, width]) => {
            conformitySheet.getColumn(colLetter).width = width;
          });
        }

        const commentsJsonPath = path.join(subfolderPath, config.excel.NOMBRE_ARCHIVO_COMENTARIOS);
        let originalSubfolderComments = {};
        try {
          const commentsContent = await fs.readFile(commentsJsonPath, 'utf-8');
          originalSubfolderComments = JSON.parse(commentsContent);
        } catch (err) {
          logger.warn(`Could not load or parse ${commentsJsonPath}: ${err.message}`);
        }

        const commentsToUse = config.services.N8N_WEBHOOK_URL
          ? await correctCommentsViaN8N(originalSubfolderComments)
          : originalSubfolderComments;

        // *** LÓGICA MODIFICADA PARA ORDEN DE FOTOS SEGÚN JSON ***
        // Tomar todas las claves del JSON de comentarios.
        // Para claves no numéricas, Object.keys() usualmente respeta el orden de inserción.
        const photoKeys = Object.keys(commentsToUse);
        // Ya NO se aplica .sort() para mantener el orden del JSON.

        const photosWithComments = [];
        logger.info(`Found ${photoKeys.length} photo keys (free names allowed) for '${subfolderName}'. Order will be based on JSON key order.`);

        for (const photoKeyStr of photoKeys) {
          const commentText = commentsToUse[photoKeyStr] || "";
          let foundPhotoPath = null;

          for (const ext of config.excel.PHOTO_CONFIG.POSSIBLE_IMAGE_EXTENSIONS) {
            const photoFileName = `${photoKeyStr}${ext.toLowerCase()}`;
            const tentativePhotoPath = path.join(subfolderPath, photoFileName);
            try {
              await fs.access(tentativePhotoPath);
              foundPhotoPath = tentativePhotoPath;
              logger.debug(`Found image: ${tentativePhotoPath}`);
              break;
            } catch { /* Archivo no encontrado, intentar siguiente extensión */ }
          }

          if (foundPhotoPath) {
            photosWithComments.push({ path: foundPhotoPath, comment: commentText });
          } else {
            logger.warn(`Image matching key '${photoKeyStr}' (with any common extension) not found in '${subfolderPath}'.`);
          }
        }
        // *** FIN DE LA LÓGICA MODIFICADA ***
        logger.info(`Processed ${photosWithComments.length} photos with comments for '${subfolderName}'.`);

        const photosPerInstance = config.excel.PHOTO_CONFIG.PHOTOS_PER_INSTANCE_STRUCTURE;
        const numInstancesRequired = photosWithComments.length > 0
          ? Math.ceil(photosWithComments.length / photosPerInstance)
          : 1;

        logger.info(`Total valid photos for '${subfolderName}': ${photosWithComments.length}. Template instances: ${numInstancesRequired}`);

        for (let i = 0; i < numInstancesRequired; i++) {
          const rowOffset = i * config.excel.CONFORMIDAD_SHEET_CONFIG.filas_por_bloque_plantilla;
          const photosForThisInstance = photosWithComments.slice(
            i * photosPerInstance,
            (i + 1) * photosPerInstance
          );
          await createConformitySheetInstance(
            conformitySheet,
            rowOffset,
            generalData,
            conformitySheetTitleTemplate,
            photosForThisInstance
          );
        }
      } else if (subfolderName.toLowerCase() !== "fotos") {
        logger.info(`Subfolder '${subfolderName}' ignored (not in FOLDER_TO_TITLE_MAP or a designated special folder).`);
      }
    }
  } catch (error) {
    logger.error('Error processing project subfolders:', error);
    return { success: false, message: `Error reading project directory: ${error.message}`, filePath: null };
  }

  if (validSubfoldersProcessedCount === 0) {
    logger.warn("No valid subfolders found to generate conformity sheets. Aborting Excel generation beyond summary sheet.");
    return { success: false, message: "No data for conformity reports found. Excel not generated.", filePath: null };
  }

  logger.info("Creating 'Mediciones OTDR' sheet...");
  const otdrSheet = workbook.addWorksheet("Mediciones OTDR", {
    views: [{ showGridLines: false }]
  });
  await createSheetMedicionesOTDR(otdrSheet, generalData, config.excel.RUTA_LOGO_GTD);
  logger.info("'Mediciones OTDR' created.");

  const projectCode = generalData["N° PROY/ COD: AX"];
  let outputFileName;
  if (projectCode && projectCode !== "SIN_CODIGO_PROYECTO") {
    outputFileName = `${sanitizeFileName(projectCode)}-ACTA DE CONFORMIDAD.xlsx`;
  } else {
    const projectFolderName = path.basename(mainProjectFolder);
    outputFileName = `${sanitizeFileName(projectFolderName)}_Consolidado_Actas.xlsx`;
    logger.warn(`'N° PROY/ COD: AX' not found or invalid in datos.json. Using folder name for Excel file: ${outputFileName}`);
  }
  const fullOutputFilePath = path.join(mainProjectFolder, outputFileName);

  try {
    await workbook.xlsx.writeFile(fullOutputFilePath);
    logger.info(`Excel file generated successfully at: ${fullOutputFilePath}`);
    return { success: true, message: "Excel generated successfully.", filePath: fullOutputFilePath };
  } catch (error) {
    logger.error(`Error saving Excel file to ${fullOutputFilePath}: ${error.message}`, error);
    return { success: false, message: `Error saving Excel: ${error.message}`, filePath: null };
  }
}

export { generateFullExcelForPath };