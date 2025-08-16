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

async function generateFullExcelForPath(requestBody) {
  const { uri: projectPath, projectDetails, services } = requestBody;
  const generalData = projectDetails || {};

  logger.info(`Starting Excel generation for project in: ${projectPath}`);

  if (!projectPath) {
    throw new Error('projectPath is undefined. Check the incoming JSON body for the "uri" key.');
  }

  const workbook = new ExcelJS.Workbook();

  const batchOfComments = {};
  if (services && Array.isArray(services)) {
    for (const service of services) {
      const serviceName = service.name;
      if (!serviceName || !config.excel.FOLDER_TO_TITLE_MAP[serviceName]) continue;
      
      batchOfComments[serviceName] = {};
      (service.photos || []).forEach(photo => {
        batchOfComments[serviceName][photo.fileName] = photo.comment;
      });
      (service.folders || []).forEach(folder => {
        (folder.photos || []).forEach(photo => {
          const uniquePhotoId = `${folder.name}/${photo.fileName}`;
          batchOfComments[serviceName][uniquePhotoId] = photo.comment;
        });
      });
    }
  }
  
  const allCorrectedComments = await correctAllCommentsInBatch(batchOfComments);
  
  await createSheetActaResumen(workbook.addWorksheet("Acta Resumen Pext", { views: [{ showGridLines: false }] }), generalData, config.excel.RUTA_LOGO_GTD);
  
  const allPossibleServices = Object.keys(config.excel.FOLDER_TO_TITLE_MAP);

  for (const serviceName of allPossibleServices) {
    const conformitySheetTitle = config.excel.FOLDER_TO_TITLE_MAP[serviceName];
    const serviceData = services.find(s => s.name === serviceName);
    
    logger.info(`Processing sheet for service: ${serviceName}`);
    const sheetName = serviceName.substring(0, 31);
    const sheet = workbook.addWorksheet(sheetName, { views: [{ showGridLines: false }] });
    
    Object.entries(config.excel.CONFORMIDAD_SHEET_CONFIG.anchos_columnas_char).forEach(([col, width]) => sheet.getColumn(col).width = width);

    let photosForExcel = [];
    if (serviceData) {
      const allPhotosData = [
        ...(serviceData.photos || []).map(p => ({ ...p, subfolderName: null })),
        ...(serviceData.folders || []).flatMap(f => f.photos.map(p => ({ ...p, subfolderName: f.name })))
      ];

      const correctedServiceComments = allCorrectedComments[serviceName] || {};

      for (const photoData of allPhotosData) {
        const uniqueId = photoData.subfolderName ? `${photoData.subfolderName}/${photoData.fileName}` : photoData.fileName;
        const correctedComment = correctedServiceComments[uniqueId] || photoData.comment;
        const finalComment = photoData.subfolderName ? `${photoData.subfolderName} - ${correctedComment}` : correctedComment;
        
        // [CODESENSEI] INICIO DE LA CORRECCIÓN: Bucle para buscar la extensión correcta.
        let foundImagePath = null;
        for (const ext of config.excel.PHOTO_CONFIG.POSSIBLE_IMAGE_EXTENSIONS) {
          const basePath = path.join(projectPath, serviceName, photoData.subfolderName || '');
          const potentialPath = path.join(basePath, `${photoData.fileName}${ext}`);
          
          try {
            await fs.access(potentialPath);
            foundImagePath = potentialPath;
            break; // Salimos del bucle en cuanto encontramos una imagen
          } catch (e) {
            // El archivo no existe con esta extensión, continuamos con la siguiente
          }
        }
        
        if (foundImagePath) {
          photosForExcel.push({ path: foundImagePath, comment: finalComment });
        } else {
          logger.warn(`Image file not found for base name: ${path.join(serviceName, photoData.subfolderName || '', photoData.fileName)}`);
        }
        // [CODESENSEI] FIN DE LA CORRECCIÓN
      }
    }
    
    const photosPerInstance = config.excel.PHOTO_CONFIG.PHOTOS_PER_INSTANCE_STRUCTURE;
    const numInstances = photosForExcel.length > 0 ? Math.ceil(photosForExcel.length / photosPerInstance) : 1;
    
    for (let i = 0; i < numInstances; i++) {
      const rowOffset = i * config.excel.CONFORMIDAD_SHEET_CONFIG.filas_por_bloque_plantilla;
      const photosForThisInstance = photosForExcel.slice(i * photosPerInstance, (i + 1) * photosPerInstance);
      await createConformitySheetInstance(sheet, rowOffset, generalData, conformitySheetTitle, photosForThisInstance);
    }
  }

  await createSheetMedicionesOTDR(workbook.addWorksheet("Mediciones OTDR", { views: [{ showGridLines: false }] }), generalData, config.excel.RUTA_LOGO_GTD);

  const projectCode = generalData.code;
  const outputFileName = projectCode
    ? `${sanitizeFileName(projectCode)}-ACTA DE CONFORMIDAD.xlsx`
    : `${sanitizeFileName(path.basename(projectPath))}_Consolidado_Actas.xlsx`;
  
  const fullOutputFilePath = path.join(projectPath, outputFileName);
  await workbook.xlsx.writeFile(fullOutputFilePath);
  logger.info(`Excel file generated successfully at: ${fullOutputFilePath}`);
  return { success: true, message: "Excel generated successfully.", filePath: fullOutputFilePath };
}

export { generateFullExcelForPath };