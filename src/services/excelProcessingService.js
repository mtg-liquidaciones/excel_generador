// [CODESENSEI] Versión refactorizada para crear todas las hojas de servicio, existan o no las carpetas.

import ExcelJS from 'exceljs';
import path from 'path';
import fs from 'fs/promises';
import config from '../config/index.js';
import { createSheetActaResumen } from '../excelGeneration/sheetActaResumen.js';
import { createSheetMedicionesOTDR } from '../excelGeneration/sheetMedicionesOTDR.js';
import { createConformitySheetInstance } from '../excelGeneration/sheetConformidad.js';
import { correctAllCommentsInBatch } from './n8nService.js';
import logger from '../utils/logger.js';

function sanitizeFileName(fileName) {
  if (!fileName) return 'default_filename';
  return fileName.replace(/[^a-z0-9_.-]+/gi, '_');
}

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
  
  // --- INICIO DE LA LÓGICA DE BATCH (se mantiene igual) ---
  const allPossibleServices = Object.keys(config.excel.FOLDER_TO_TITLE_MAP);
  const batchOfComments = {};
  for (const subfolderName of allPossibleServices) {
    const commentsJsonPath = path.join(mainProjectFolder, subfolderName, config.excel.NOMBRE_ARCHIVO_COMENTARIOS);
    try {
      const commentsContent = await fs.readFile(commentsJsonPath, 'utf-8');
      batchOfComments[subfolderName] = JSON.parse(commentsContent);
    } catch (err) {
      // Si el archivo de comentarios no existe, simplemente no lo añadimos al lote
      if (err.code !== 'ENOENT') {
        logger.warn(`Could not load or parse comments for '${subfolderName}'.`);
      }
    }
  }
  
  logger.info(`Sending a batch of comments from ${Object.keys(batchOfComments).length} subfolders to n8n.`);
  const allCorrectedComments = await correctAllCommentsInBatch(batchOfComments);
  // --- FIN DE LA LÓGICA DE BATCH ---


  // --- CREACIÓN DE HOJAS ---
  logger.info("Creating 'Acta Resumen Pext' sheet...");
  const resumenSheet = workbook.addWorksheet("Acta Resumen Pext", { views: [{ showGridLines: false }] });
  await createSheetActaResumen(resumenSheet, generalData, config.excel.RUTA_LOGO_GTD);
  logger.info("'Acta Resumen Pext' created.");

  // [CODESENSEI] INICIO DE LA LÓGICA MODIFICADA
  // Iteramos sobre la lista DEFINIDA de servicios en la configuración, no sobre las carpetas que encontramos.
  logger.info("Creating conformity sheets for all defined services...");

  for (const subfolderName of allPossibleServices) {
    const subfolderPath = path.join(mainProjectFolder, subfolderName);
    const conformitySheetTitleTemplate = config.excel.FOLDER_TO_TITLE_MAP[subfolderName];
    
    logger.info(`Processing sheet for service: ${subfolderName}`);

    const sheetName = subfolderName.substring(0, 31);
    const conformitySheet = workbook.addWorksheet(sheetName, { views: [{ showGridLines: false }] });
    
    const colWidths = config.excel.CONFORMIDAD_SHEET_CONFIG.anchos_columnas_char;
    if (colWidths) {
      Object.entries(colWidths).forEach(([col, width]) => conformitySheet.getColumn(col).width = width);
    }

    let photosWithComments = [];
    // [CODESENSEI] Verificamos si la carpeta para este servicio existe en el disco.
    const folderExists = await fs.access(subfolderPath).then(() => true).catch(() => false);

    if (folderExists) {
      logger.info(`Folder '${subfolderName}' found. Populating with data...`);
      // [CODESENSEI] Si la carpeta existe, usamos los comentarios (ya corregidos en lote) y buscamos las fotos.
      const commentsToUse = allCorrectedComments[subfolderName] || {};
      const photoKeys = Object.keys(commentsToUse);
      
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
    } else {
      // [CODESENSEI] Si la carpeta no existe, el array 'photosWithComments' se queda vacío.
      logger.warn(`Folder '${subfolderName}' not found. An empty template sheet will be created.`);
    }

    // [CODESENSEI] Esta lógica ahora se ejecuta siempre, ya sea con datos o con 'photosWithComments' vacío.
    const photosPerInstance = config.excel.PHOTO_CONFIG.PHOTOS_PER_INSTANCE_STRUCTURE;
    const numInstancesRequired = photosWithComments.length > 0 ? Math.ceil(photosWithComments.length / photosPerInstance) : 1;
    
    logger.info(`Total valid photos for '${subfolderName}': ${photosWithComments.length}. Template instances: ${numInstancesRequired}`);
    
    for (let i = 0; i < numInstancesRequired; i++) {
      const rowOffset = i * config.excel.CONFORMIDAD_SHEET_CONFIG.filas_por_bloque_plantilla;
      const photosForThisInstance = photosWithComments.slice(i * photosPerInstance, (i + 1) * photosPerInstance);
      // [CODESENSEI] Llamamos a la función que dibuja la plantilla, incluso si no hay fotos.
      await createConformitySheetInstance(conformitySheet, rowOffset, generalData, conformitySheetTitleTemplate, photosForThisInstance);
    }
  }
  // [CODESENSEI] FIN DE LA LÓGICA MODIFICADA

  logger.info("Creating 'Mediciones OTDR' sheet...");
  const otdrSheet = workbook.addWorksheet("Mediciones OTDR", { views: [{ showGridLines: false }] });
  await createSheetMedicionesOTDR(otdrSheet, generalData, config.excel.RUTA_LOGO_GTD);
  logger.info("'Mediciones OTDR' created.");

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