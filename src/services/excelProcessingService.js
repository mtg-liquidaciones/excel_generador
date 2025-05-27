// src/services/excelProcessingService.js

import ExcelJS from 'exceljs';
import path from 'path';
import fs from 'fs/promises';
import config from '../config/index.js'; // Main configuration
// Assuming a logger utility will be created at src/utils/logger.js
// import logger from '../utils/logger.js';
import { createSheetActaResumen } from '../excelGeneration/sheetActaResumen.js';
import { createSheetMedicionesOTDR } from '../excelGeneration/sheetMedicionesOTDR.js';
import { createConformitySheetInstance } from '../excelGeneration/sheetConformidad.js';
import { correctCommentsViaN8N } from './n8nService.js';

// Placeholder logger - replace with actual logger import
const logger = {
  info: (message) => console.log(`[INFO] ${message}`),
  warn: (message) => console.warn(`[WARN] ${message}`),
  error: (message, error) => console.error(`[ERROR] ${message}`, error || ''),
  debug: (message) => console.log(`[DEBUG] ${message}`),
};

/**
 * Sanitizes a string to be used as a valid filename.
 * Replaces most non-alphanumeric characters with an underscore.
 * @param {string} fileName - The original filename string.
 * @returns {string} - The sanitized filename.
 */
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

  // 1. Load main datos.json
  const mainDataJsonPath = path.join(mainProjectFolder, config.excel.NOMBRE_ARCHIVO_DATOS_PRINCIPAL);
  let generalData = {};
  try {
    const fileContent = await fs.readFile(mainDataJsonPath, 'utf-8');
    generalData = JSON.parse(fileContent);
    logger.info(`Main datos.json loaded from ${mainDataJsonPath}.`);
  } catch (error) {
    logger.error(`Error loading or parsing ${mainDataJsonPath}: ${error.message}`, error);
    // Python script continues with empty datos_generales, so we do the same.
    // Consider if this should be a hard failure depending on requirements.
  }

  // 2. Initialize Workbook
  const workbook = new ExcelJS.Workbook();

  // 3. Create "Acta Resumen Pext" sheet
  logger.info("Creating 'Acta Resumen Pext' sheet...");
  const resumenSheet = workbook.addWorksheet("Acta Resumen Pext", {
    views: [{ showGridLines: false }]
  });
  // The sheet creation functions are expected to be async if they perform async operations (e.g., image handling)
  await createSheetActaResumen(resumenSheet, generalData, config.excel.RUTA_LOGO_GTD);
  logger.info("'Acta Resumen Pext' created.");

  let validSubfoldersProcessedCount = 0;

  // 4. Process subfolders for "Actas de Conformidad"
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

        const sheetName = subfolderName.substring(0, 31); // Excel sheet name limit
        const conformitySheet = workbook.addWorksheet(sheetName, {
          views: [{ showGridLines: false }]
        });

        // Apply column widths from config
        const colWidths = config.excel.CONFORMIDAD_SHEET_CONFIG.anchos_columnas_char;
        if (colWidths) {
          Object.entries(colWidths).forEach(([colLetter, width]) => {
            conformitySheet.getColumn(colLetter).width = width;
          });
        }

        // Load comentarios.json for the subfolder
        const commentsJsonPath = path.join(subfolderPath, config.excel.NOMBRE_ARCHIVO_COMENTARIOS);
        let originalSubfolderComments = {};
        try {
          const commentsContent = await fs.readFile(commentsJsonPath, 'utf-8');
          originalSubfolderComments = JSON.parse(commentsContent);
        } catch (err) {
          logger.warn(`Could not load or parse ${commentsJsonPath}: ${err.message}`);
        }

        // Correct comments via N8N (if URL is configured)
        const commentsToUse = config.services.N8N_WEBHOOK_URL
          ? await correctCommentsViaN8N(originalSubfolderComments) // N8N URL is now read from config inside n8nService
          : originalSubfolderComments;

        // Find and prepare photos with their comments
        const photosWithComments = [];
        const photoKeys = Object.keys(commentsToUse)
          .filter(key => /^\d+$/.test(key)) // Numeric keys for photos
          .sort((a, b) => parseInt(a, 10) - parseInt(b, 10));

        for (const photoKey of photoKeys) {
          const commentText = commentsToUse[photoKey] || "";
          let foundPhotoPath = null;
          for (const ext of config.excel.PHOTO_CONFIG.POSSIBLE_IMAGE_EXTENSIONS) {
            const photoFileName = `${photoKey}${ext.toLowerCase()}`;
            const tentativePhotoPath = path.join(subfolderPath, photoFileName);
            try {
              await fs.access(tentativePhotoPath); // Check file existence
              foundPhotoPath = tentativePhotoPath;
              break;
            } catch { /* File not found, try next extension */ }
          }
          if (foundPhotoPath) {
            photosWithComments.push({ path: foundPhotoPath, comment: commentText });
          } else {
            logger.warn(`Image '${photoKey}.*' not found in '${subfolderPath}'.`);
          }
        }
        logger.info(`Found ${photosWithComments.length} photos for '${subfolderName}'.`);

        const photosPerInstance = config.excel.PHOTO_CONFIG.PHOTOS_PER_INSTANCE_STRUCTURE;
        const numInstancesRequired = photosWithComments.length > 0
          ? Math.ceil(photosWithComments.length / photosPerInstance)
          : 1; // Always one instance, even if no photos (as per Python logic)

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


  // 5. Check if any conformity sheets were actually created. If not, maybe abort.
  // Python: if not libro_excel.sheetnames and subcarpetas_procesadas_validas == 0:
  // (This check was a bit off as Resumen was always created first)
  // More accurate: if no conformity sheets, don't proceed to OTDR or save.
  if (validSubfoldersProcessedCount === 0) {
    logger.warn("No valid subfolders found to generate conformity sheets. Aborting Excel generation beyond summary sheet.");
    // Depending on requirements, you might save an Excel with only the summary,
    // or return an error indicating no main content was processed.
    // The Python code implies not saving and not making OTDR.
    return { success: false, message: "No data for conformity reports found. Excel not generated.", filePath: null };
  }

  // 6. Create "Mediciones OTDR" sheet
  logger.info("Creating 'Mediciones OTDR' sheet...");
  const otdrSheet = workbook.addWorksheet("Mediciones OTDR", {
    views: [{ showGridLines: false }]
  });
  await createSheetMedicionesOTDR(otdrSheet, generalData, config.excel.RUTA_LOGO_GTD);
  logger.info("'Mediciones OTDR' created.");

  // 7. Determine Output Filename
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

  // 8. Save Workbook
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