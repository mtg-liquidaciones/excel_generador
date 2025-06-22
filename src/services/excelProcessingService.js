// [CODESENSEI] Versión refactorizada para procesamiento por lotes (batch processing)

import ExcelJS from 'exceljs';
import path from 'path';
import fs from 'fs/promises';
import config from '../config/index.js';
import { createSheetActaResumen } from '../excelGeneration/sheetActaResumen.js';
import { createSheetMedicionesOTDR } from '../excelGeneration/sheetMedicionesOTDR.js';
import { createConformitySheetInstance } from '../excelGeneration/sheetConformidad.js';
import { correctAllCommentsInBatch } from './n8nService.js'; // [CODESENSEI] Importamos la nueva función de batch
import logger from '../utils/logger.js';

// ... (la función sanitizeFileName se mantiene igual)
function sanitizeFileName(fileName) {
  if (!fileName) return 'default_filename';
  return fileName.replace(/[^a-z0-9_.-]+/gi, '_');
}


async function generateFullExcelForPath(mainProjectFolder) {
  logger.info(`Starting Excel generation for project in: ${mainProjectFolder}`);

  // --- LECTURA INICIAL DE DATOS ---
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
  const projectContents = await fs.readdir(mainProjectFolder, { withFileTypes: true });
  const subfoldersToProcess = projectContents.filter(dirent => dirent.isDirectory() && config.excel.FOLDER_TO_TITLE_MAP[dirent.name]);

  if (subfoldersToProcess.length === 0) {
    logger.warn("No valid subfolders found to generate conformity sheets. Aborting.");
    return { success: false, message: "No data for conformity reports found.", filePath: null };
  }
  
  // --- [CODESENSEI] INICIO DE LA LÓGICA DE BATCH ---

  // 1. Recolectar todos los comentarios en un solo objeto
  const batchOfComments = {};
  for (const dirent of subfoldersToProcess) {
    const subfolderName = dirent.name;
    const commentsJsonPath = path.join(mainProjectFolder, subfolderName, config.excel.NOMBRE_ARCHIVO_COMENTARIOS);
    try {
      const commentsContent = await fs.readFile(commentsJsonPath, 'utf-8');
      batchOfComments[subfolderName] = JSON.parse(commentsContent);
    } catch (err) {
      logger.warn(`Could not load or parse comments for '${subfolderName}'. It will be skipped in n8n call.`);
      batchOfComments[subfolderName] = {}; // Añadir objeto vacío para mantener la estructura
    }
  }

  // 2. Hacer UNA SOLA llamada a n8n con todo el lote de comentarios
  logger.info(`Sending a batch of comments from ${Object.keys(batchOfComments).length} subfolders to n8n.`);
  const allCorrectedComments = await correctAllCommentsInBatch(batchOfComments);

  // --- [CODESENSEI] FIN DE LA LÓGICA DE BATCH ---

  // 3. Crear las hojas del Excel usando los datos ya procesados
  logger.info("Creating 'Acta Resumen Pext' sheet...");
  const resumenSheet = workbook.addWorksheet("Acta Resumen Pext", { views: [{ showGridLines: false }] });
  await createSheetActaResumen(resumenSheet, generalData, config.excel.RUTA_LOGO_GTD);
  logger.info("'Acta Resumen Pext' created.");

  for (const dirent of subfoldersToProcess) {
    const subfolderName = dirent.name;
    const subfolderPath = path.join(mainProjectFolder, subfolderName);
    const conformitySheetTitleTemplate = config.excel.FOLDER_TO_TITLE_MAP[subfolderName];
    
    // [CODESENSEI] Obtenemos los comentarios ya corregidos de la respuesta en lote.
    const commentsToUse = allCorrectedComments[subfolderName] || {};
    
    // El resto de la lógica para encontrar fotos y generar la hoja se mantiene, pero ahora es más rápido y robusto.
    // ... (la lógica de búsqueda de fotos insensible a mayúsculas que ya implementamos) ...
    const photoKeys = Object.keys(commentsToUse);
    const photosWithComments = [];
    
    if (photoKeys.length > 0) {
        const filesInSubfolder = await fs.readdir(subfolderPath);
        const filesInSubfolderLower = filesInSubfolder.map(f => f.toLowerCase());

        for (const photoKeyStr of photoKeys) {
            const commentText = commentsToUse[photoKeyStr] || "";
            let foundPhotoPath = null;

            for (const ext of config.excel.PHOTO_CONFIG.POSSIBLE_IMAGE_EXTENSIONS) {
                const targetFileNameLower = `${photoKeyStr}${ext}`.toLowerCase();
                const fileIndex = filesInSubfolderLower.indexOf(targetFileNameLower);

                if (fileIndex !== -1) {
                    foundPhotoPath = path.join(subfolderPath, filesInSubfolder[fileIndex]);
                    break;
                }
            }
            if (foundPhotoPath) {
                photosWithComments.push({ path: foundPhotoPath, comment: commentText });
            } else {
                logger.warn(`Image matching key '${photoKeyStr}' not found in '${subfolderName}'.`);
            }
        }
    }
    
    if (photosWithComments.length === 0) {
        logger.warn(`No images were matched for '${subfolderName}'. A blank report sheet will be created.`);
    }

    const sheetName = subfolderName.substring(0, 31);
    const conformitySheet = workbook.addWorksheet(sheetName, { views: [{ showGridLines: false }] });
    const colWidths = config.excel.CONFORMIDAD_SHEET_CONFIG.anchos_columnas_char;
    if (colWidths) {
        Object.entries(colWidths).forEach(([col, width]) => conformitySheet.getColumn(col).width = width);
    }
    
    const photosPerInstance = config.excel.PHOTO_CONFIG.PHOTOS_PER_INSTANCE_STRUCTURE;
    const numInstancesRequired = photosWithComments.length > 0 ? Math.ceil(photosWithComments.length / photosPerInstance) : 1;
    
    for (let i = 0; i < numInstancesRequired; i++) {
        // ... (lógica para crear la instancia de la hoja de conformidad) ...
        const rowOffset = i * config.excel.CONFORMIDAD_SHEET_CONFIG.filas_por_bloque_plantilla;
        const photosForThisInstance = photosWithComments.slice(i * photosPerInstance, (i + 1) * photosPerInstance);
        await createConformitySheetInstance(conformitySheet, rowOffset, generalData, conformitySheetTitleTemplate, photosForThisInstance);
    }
  }

  // --- CREACIÓN DE HOJAS FINALES Y GUARDADO ---
  logger.info("Creating 'Mediciones OTDR' sheet...");
  const otdrSheet = workbook.addWorksheet("Mediciones OTDR", { views: [{ showGridLines: false }] });
  await createSheetMedicionesOTDR(otdrSheet, generalData, config.excel.RUTA_LOGO_GTD);
  logger.info("'Mediciones OTDR' created.");

  // ... (la lógica para generar el nombre del archivo y guardarlo se mantiene igual) ...
  const projectCode = generalData["N° PROY/ COD: AX"];
  let outputFileName;
  if (projectCode && projectCode !== "SIN_CODIGO_PROYECTO") {
    outputFileName = `${sanitizeFileName(projectCode)}-ACTA DE CONFORMIDAD.xlsx`;
  } else {
    const projectFolderName = path.basename(mainProjectFolder);
    outputFileName = `${sanitizeFileName(projectFolderName)}_Consolidado_Actas.xlsx`;
  }
  const fullOutputFilePath = path.join(mainProjectFolder, outputFileName);
  await workbook.xlsx.writeFile(fullOutputFilePath);
  logger.info(`Excel file generated successfully at: ${fullOutputFilePath}`);
  return { success: true, message: "Excel generated successfully.", filePath: fullOutputFilePath };
}

export { generateFullExcelForPath };
