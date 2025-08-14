// [CODESENSEI] Versión final y corregida.

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
  // [CODESENSEI] CORRECCIÓN CLAVE: Usamos 'uri' y lo renombramos a 'projectPath'.
  const { uri: projectPath, projectDetails, services } = requestBody;
  const generalData = projectDetails || {};

  logger.info(`Starting Excel generation for project in: ${projectPath}`);

  if (!projectPath) {
    throw new Error('projectPath is undefined after destructuring. Check the incoming JSON body for the "uri" key.');
  }

  const workbook = new ExcelJS.Workbook();

  const batchOfComments = {};
  if (services && Array.isArray(services)) {
    for (const service of services) {
      const serviceName = service.name;
      if (!serviceName) continue;
      
      batchOfComments[serviceName] = {};
      if (service.photos && service.photos.length > 0) {
        for (const photo of service.photos) {
          batchOfComments[serviceName][photo.fileName] = photo.comment;
        }
      }
      if (service.folders && service.folders.length > 0) {
        for (const folder of service.folders) {
          for (const photo of folder.photos) {
            const uniquePhotoId = `${folder.name}/${photo.fileName}`;
            batchOfComments[serviceName][uniquePhotoId] = photo.comment;
          }
        }
      }
    }
  }

  logger.info(`Sending a batch of comments from ${Object.keys(batchOfComments).length} services to n8n.`);
  const allCorrectedComments = await correctAllCommentsInBatch(batchOfComments);
  
  logger.info("Creating 'Acta Resumen Pext' sheet...");
  const resumenSheet = workbook.addWorksheet("Acta Resumen Pext", { views: [{ showGridLines: false }] });
  await createSheetActaResumen(resumenSheet, generalData, config.excel.RUTA_LOGO_GTD);
  logger.info("'Acta Resumen Pext' created.");

  logger.info("Creating conformity sheets for all defined services...");
  const allPossibleServices = Object.keys(config.excel.FOLDER_TO_TITLE_MAP);

  for (const serviceName of allPossibleServices) {
    const conformitySheetTitleTemplate = config.excel.FOLDER_TO_TITLE_MAP[serviceName];
    const serviceData = services.find(s => s.name === serviceName);
    
    logger.info(`Processing sheet for service: ${serviceName}`);
    const sheetName = serviceName.substring(0, 31);
    const conformitySheet = workbook.addWorksheet(sheetName, { views: [{ showGridLines: false }] });
    
    const colWidths = config.excel.CONFORMIDAD_SHEET_CONFIG.anchos_columnas_char;
    if (colWidths) {
      Object.entries(colWidths).forEach(([col, width]) => conformitySheet.getColumn(col).width = width);
    }

    let photosForExcel = [];
    const correctedServiceComments = allCorrectedComments[serviceName] || {};
    
    if (serviceData) {
      const allPhotosData = [];
      // Fotos en la raíz
      (serviceData.photos || []).forEach(p => allPhotosData.push({ ...p, subfolderName: null, serviceName }));
      // Fotos en carpetas
      (serviceData.folders || []).forEach(f => {
        (f.photos || []).forEach(p => allPhotosData.push({ ...p, subfolderName: f.name, serviceName }));
      });
      
      for (const photoData of allPhotosData) {
        const uniqueId = photoData.subfolderName ? `${photoData.subfolderName}/${photoData.fileName}` : photoData.fileName;
        const correctedComment = correctedServiceComments[uniqueId] || photoData.comment;
        const finalComment = photoData.subfolderName ? `${photoData.subfolderName} - ${correctedComment}` : correctedComment;
        
        const imagePath = photoData.subfolderName
          ? path.join(projectPath, photoData.serviceName, photoData.subfolderName, `${photoData.fileName}.jpg`)
          : path.join(projectPath, photoData.serviceName, `${photoData.fileName}.jpg`);
        
        const imageExists = await fs.access(imagePath).then(() => true).catch(() => false);
        if (imageExists) {
            photosForExcel.push({ path: imagePath, comment: finalComment });
        } else {
            logger.warn(`Image file not found at path: ${imagePath}`);
        }
      }
    }
    
    const photosPerInstance = config.excel.PHOTO_CONFIG.PHOTOS_PER_INSTANCE_STRUCTURE;
    const numInstancesRequired = photosForExcel.length > 0 ? Math.ceil(photosForExcel.length / photosPerInstance) : 1;
    
    for (let i = 0; i < numInstancesRequired; i++) {
      const rowOffset = i * config.excel.CONFORMIDAD_SHEET_CONFIG.filas_por_bloque_plantilla;
      const photosForThisInstance = photosForExcel.slice(i * photosPerInstance, (i + 1) * photosPerInstance);
      await createConformitySheetInstance(conformitySheet, rowOffset, generalData, conformitySheetTitleTemplate, photosForThisInstance);
    }
  }

  logger.info("Creating 'Mediciones OTDR' sheet...");
  const otdrSheet = workbook.addWorksheet("Mediciones OTDR", { views: [{ showGridLines: false }] });
  await createSheetMedicionesOTDR(otdrSheet, generalData, config.excel.RUTA_LOGO_GTD);
  logger.info("'Mediciones OTDR' created.");

  const projectCode = generalData.code;
  let outputFileName;
  if (projectCode && projectCode !== "SIN_CODIGO_PROYECTO") {
    outputFileName = `${sanitizeFileName(projectCode)}-ACTA DE CONFORMIDAD.xlsx`;
  } else {
    const projectFolderName = path.basename(projectPath);
    outputFileName = `${sanitizeFileName(projectFolderName)}_Consolidado_Actas.xlsx`;
  }
  
  const fullOutputFilePath = path.join(projectPath, outputFileName);
  await workbook.xlsx.writeFile(fullOutputFilePath);
  logger.info(`Excel file generated successfully at: ${fullOutputFilePath}`);
  return { success: true, message: "Excel generated successfully.", filePath: fullOutputFilePath };
}

export { generateFullExcelForPath };