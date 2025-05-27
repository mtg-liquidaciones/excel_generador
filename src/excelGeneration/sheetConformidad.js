// src/excelGeneration/sheetConformidad.js

import config from '../config/index.js';
import {
  applyCellStyles,
  applyOuterBorder,
  insertResizedImage,
  offsetCellOrRangeRef,
  parseCellRef, // Usada por insertResizedImage y offsetCellOrRangeRef
  // getColumnNumber // Usada por parseCellRef, se asume que está en excelUtils o parseCellRef la maneja
} from './excelUtils.js';

// Placeholder logger - Reemplázalo con tu logger configurado
const logger = {
  info: (message) => console.log(`[INFO] sheetConformidad: ${message}`),
  warn: (message) => console.warn(`[WARN] sheetConformidad: ${message}`),
  error: (message, errorDetails) => {
    const detailsString = errorDetails && typeof errorDetails === 'object' ? JSON.stringify(errorDetails) : errorDetails || '';
    console.error(`[ERROR] sheetConformidad: ${message}`, detailsString);
  },
  debug: (message) => console.log(`[DEBUG] sheetConformidad: ${message}`),
};

// --- Funciones Auxiliares para la Plantilla de Conformidad ---

/**
 * Aplica las alturas de fila base definidas en la configuración.
 * MODIFICADO para no sobrescribir la fila 114 (o equivalente) del bloque anterior.
 * @param {import('exceljs').Worksheet} sheet
 * @param {number} rowOffset
 * @param {object} conformityConfig
 */
function _applyRowHeightsConformity(sheet, rowOffset, conformityConfig) {
  if (conformityConfig.alturas_filas_base) {
    for (const baseRowStr in conformityConfig.alturas_filas_base) {
      const baseRow = parseInt(baseRowStr, 10);
      const height = conformityConfig.alturas_filas_base[baseRow];
      const actualRowNumber = baseRow + rowOffset;

      // Si este NO es el primer bloque (rowOffset > 0) Y estamos
      // a punto de establecer la altura para la baseRow 1 del bloque ACTUAL:
      // esta fila (actualRowNumber) era la fila (1 + filas_por_bloque_plantilla) del bloque ANTERIOR.
      // Específicamente, si filas_por_bloque_plantilla es 113, la fila 114 del bloque anterior
      // es la fila (1 + offset_del_bloque_anterior + 113) = (1 + rowOffset).
      // No la sobrescribimos para preservar su altura personalizada de 28.
      if (rowOffset > 0 && baseRow === 1) {
        // Esta condición se cumple cuando estamos procesando la baseRow 1 del segundo bloque en adelante.
        // La fila actual (1 + rowOffset) fue la fila (filas_por_bloque_plantilla + 1) del bloque anterior.
        // Ejemplo: Si filas_por_bloque_plantilla = 113.
        // Para el bloque 0 (rowOffset=0), la fila 114 se establece en 28.
        // Para el bloque 1 (rowOffset=113), esta función intentaría establecer la altura de la fila (1+113)=114.
        // Debemos saltar esto para que la fila 114 conserve su altura de 28.
        logger.debug(`_applyRowHeightsConformity: Skipping height set for base row 1 (actual row ${actualRowNumber}) of block at offset ${rowOffset} to preserve previous block's custom height.`);
        continue; // Saltar la aplicación de altura para esta fila específica
      }

      sheet.getRow(actualRowNumber).height = height;
    }
  }
}

/**
 * Aplica los textos fijos definidos en la configuración.
 * @param {import('exceljs').Worksheet} sheet
 * @param {number} rowOffset
 * @param {object} conformityConfig
 * @param {object} commonFonts
 * @param {object} commonAlignments
 */
function _applyFixedTextsConformity(sheet, rowOffset, conformityConfig, commonFonts, commonAlignments) {
  if (conformityConfig.textos_fijos_base) {
    for (const baseCellRef in conformityConfig.textos_fijos_base) {
      const data = conformityConfig.textos_fijos_base[baseCellRef];
      const actualCellRef = offsetCellOrRangeRef(baseCellRef, rowOffset);
      const cell = sheet.getCell(actualCellRef);

      const style = { value: data.texto };
      if (data.font_key && commonFonts[data.font_key]) {
        style.font = commonFonts[data.font_key];
      }
      if (data.alignment_key && commonAlignments[data.alignment_key]) {
        style.alignment = commonAlignments[data.alignment_key];
      }
      applyCellStyles(cell, style);
    }
  }
}

/**
 * Combina las celdas según la configuración base.
 * @param {import('exceljs').Worksheet} sheet
 * @param {number} rowOffset
 * @param {object} conformityConfig
 */
function _mergeCellsConformity(sheet, rowOffset, conformityConfig) {
  if (conformityConfig.celdas_a_combinar_base) {
    conformityConfig.celdas_a_combinar_base.forEach(baseRange => {
      const actualRange = offsetCellOrRangeRef(baseRange, rowOffset);
      try {
        sheet.mergeCells(actualRange);
      } catch (e) {
        logger.error(`Error merging cells '${actualRange}': ${e.message}`);
      }
    });
  }
}

/**
 * Crea las secciones para fotos y comentarios.
 * @param {import('exceljs').Worksheet} sheet
 * @param {number} rowOffset
 * @param {Array<{path: string, comment: string}>} photosForInstance
 * @param {object} conformityConfig
 * @param {object} commonStyles - Contiene COMMON_FONTS, COMMON_ALIGNMENTS, BORDER_STYLES
 */
async function _createPhotoCommentSections(sheet, rowOffset, photosForInstance, conformityConfig, commonStyles) {
  const basePhotoSectionStartRow = 21;
  const absolutePhotoSectionStartRow = basePhotoSectionStartRow + rowOffset;

  const photoSlotsLayout = [
    { photoAnchorCol: 'D', commentAnchorCol: 'D', photoBlockRowOffset: 0 },
    { photoAnchorCol: 'S', commentAnchorCol: 'S', photoBlockRowOffset: 0 },
    { photoAnchorCol: 'D', commentAnchorCol: 'D', photoBlockRowOffset: 26 },
    { photoAnchorCol: 'S', commentAnchorCol: 'S', photoBlockRowOffset: 26 },
    { photoAnchorCol: 'D', commentAnchorCol: 'D', photoBlockRowOffset: 52 },
    { photoAnchorCol: 'S', commentAnchorCol: 'S', photoBlockRowOffset: 52 },
  ];
  
  const photoAreaRelStartRow = 0; 
  const photoAreaHeightRows = 22; 
  const commentAreaRelStartRow = photoAreaHeightRows;
  const commentAreaHeightRows = 2;  

  for (let blockIndex = 0; blockIndex < 3; blockIndex++) {
    const currentBlockBaseRow = absolutePhotoSectionStartRow + (blockIndex * 26);

    const photoAreaStartActualRow = currentBlockBaseRow + photoAreaRelStartRow;
    const photoAreaEndActualRow = photoAreaStartActualRow + photoAreaHeightRows - 1;
    for (let r = photoAreaStartActualRow; r <= photoAreaEndActualRow; r++) {
      sheet.getRow(r).height = 12.75;
      sheet.mergeCells(`B${r}:C${r}`);
    }
    applyOuterBorder(sheet, `D${photoAreaStartActualRow}:P${photoAreaEndActualRow}`, commonStyles.BORDER_STYLES.THIN_SIDE);
    applyOuterBorder(sheet, `S${photoAreaStartActualRow}:AE${photoAreaEndActualRow}`, commonStyles.BORDER_STYLES.THIN_SIDE);

    const commentAreaStartActualRow = currentBlockBaseRow + commentAreaRelStartRow;
    const commentAreaEndActualRow = commentAreaStartActualRow + commentAreaHeightRows - 1;
    for (let r = commentAreaStartActualRow; r <= commentAreaEndActualRow; r++) {
      sheet.getRow(r).height = 16.75;
      sheet.mergeCells(`B${r}:C${r}`);
    }
    sheet.mergeCells(`D${commentAreaStartActualRow}:P${commentAreaEndActualRow}`);
    sheet.mergeCells(`S${commentAreaStartActualRow}:AE${commentAreaEndActualRow}`);
    applyOuterBorder(sheet, `D${commentAreaStartActualRow}:P${commentAreaEndActualRow}`, commonStyles.BORDER_STYLES.THIN_SIDE);
    applyOuterBorder(sheet, `S${commentAreaStartActualRow}:AE${commentAreaEndActualRow}`, commonStyles.BORDER_STYLES.THIN_SIDE);
    
    const spacingStartRow = commentAreaEndActualRow + 1;
    const spacingEndRow = spacingStartRow + 1; 
    for (let r = spacingStartRow; r <= spacingEndRow; r++) {
      if (sheet.getRow(r).height === undefined || sheet.getRow(r).height < 16.75 ) {
         sheet.getRow(r).height = 16.75;
      }
    }
  }

  for (let i = 0; i < photosForInstance.length && i < photoSlotsLayout.length; i++) {
    const photoInfo = photosForInstance[i];
    const slot = photoSlotsLayout[i];
    const photoAnchorCell = `${slot.photoAnchorCol}${absolutePhotoSectionStartRow + slot.photoBlockRowOffset + photoAreaRelStartRow}`;
    const commentAnchorCell = `${slot.commentAnchorCol}${absolutePhotoSectionStartRow + slot.photoBlockRowOffset + commentAreaRelStartRow}`;

    if (photoInfo && photoInfo.path) {
      try {
        await insertResizedImage(
          sheet, photoInfo.path, photoAnchorCell,
          conformityConfig.CONFORMIDAD_PHOTO_PIXEL_WIDTH,
          conformityConfig.CONFORMIDAD_PHOTO_PIXEL_HEIGHT,
          false
        );
      } catch (e) {
        logger.error(`Failed to insert image ${photoInfo.path} at ${photoAnchorCell}: ${e.message}`);
      }
    }

    if (photoInfo && photoInfo.comment) {
      applyCellStyles(sheet.getCell(commentAnchorCell), {
        value: photoInfo.comment,
        font: commonStyles.COMMON_FONTS.comment_arial_11_center,
        alignment: commonStyles.COMMON_ALIGNMENTS.center_center_wrap
      });
    }
  }
}

/**
 * Aplica bordes específicos para la plantilla de conformidad.
 * @param {import('exceljs').Worksheet} sheet
 * @param {number} rowOffset
 * @param {object} conformityConfig
 * @param {object} borderStyles
 */
function _applySpecificBordersConformity(sheet, rowOffset, conformityConfig, borderStyles) {
  const dataFieldRangesForThinBorder = [
    'T8:AE8', 'E10:P10', 'T10:Y10', 'Z10:AE10',
    'E12:P12', 'T12:Y12', 'Z12:AE12', 'E14:P14',
    'T14:Y14', 'Z14:AE14', 'H16:AE16', 'E18:AE18',
  ];
  dataFieldRangesForThinBorder.forEach(baseRange => {
    const actualRange = offsetCellOrRangeRef(baseRange, rowOffset);
    const [startCellRef] = actualRange.split(':');
    const cell = sheet.getCell(startCellRef);
    cell.border = borderStyles.BORDER_THIN_ALL_SIDES;
  });

  const thinBottomBorder = { bottom: borderStyles.THIN_SIDE };
  sheet.getCell(offsetCellOrRangeRef('H99', rowOffset)).border = thinBottomBorder;
  sheet.getCell(offsetCellOrRangeRef('T99', rowOffset)).border = thinBottomBorder;
  sheet.getCell(offsetCellOrRangeRef('E102', rowOffset)).border = thinBottomBorder;
  sheet.getCell(offsetCellOrRangeRef('T102', rowOffset)).border = thinBottomBorder;

  const obsRows = [105, 106, 107];
  const dottedBottomBorder = { bottom: borderStyles.DOTTED_SIDE };
  obsRows.forEach(baseRow => {
    sheet.getCell(offsetCellOrRangeRef(`D${baseRow}`, rowOffset)).border = dottedBottomBorder;
    sheet.getCell(offsetCellOrRangeRef(`Q${baseRow}`, rowOffset)).border = dottedBottomBorder;
  });

  applyOuterBorder(sheet, offsetCellOrRangeRef('D105:P107', rowOffset), borderStyles.THIN_SIDE);
  applyOuterBorder(sheet, offsetCellOrRangeRef('Q105:AE107', rowOffset), borderStyles.THIN_SIDE);

  obsRows.forEach(baseRow => {
    const r = baseRow + rowOffset;
    const cellP = sheet.getCell(`P${r}`);
    cellP.border = { ...(cellP.border || {}), right: borderStyles.THIN_SIDE };
    const cellQ = sheet.getCell(`Q${r}`);
    cellQ.border = { ...(cellQ.border || {}), left: borderStyles.THIN_SIDE };
  });
}

/**
 * Creates one instance of the conformity sheet template.
 * @param {import('exceljs').Worksheet} sheet
 * @param {number} rowOffset
 * @param {object} generalData
 * @param {string} conformitySheetTitle
 * @param {Array<{path: string, comment: string}>} photosForInstance
 */
async function createConformitySheetInstance(
  sheet,
  rowOffset,
  generalData,
  conformitySheetTitle,
  photosForInstance
) {
  const {
    CONFORMIDAD_SHEET_CONFIG: conformityConfig,
    RUTA_LOGO_GTD,
    COMMON_FONTS,
    COMMON_ALIGNMENTS,
    BORDER_STYLES,
    MAIN_JSON_CELL_MAP_CONFORMIDAD
  } = config.excel;

  // Aplicar alturas de fila base (MODIFICADO para no sobreescribir la fila 114 del bloque anterior)
  _applyRowHeightsConformity(sheet, rowOffset, conformityConfig);

  // Insertar Logo GTD usando insertResizedImage
  const logoBaseCellRef = conformityConfig.celda_logo_gtd;
  const actualLogoCellRef = offsetCellOrRangeRef(logoBaseCellRef, rowOffset);
  logger.debug(`Attempting to insert logo for Conformity Sheet (offset ${rowOffset}) via insertResizedImage. Path: "${RUTA_LOGO_GTD}", Anchor: ${actualLogoCellRef}`);
  if (!RUTA_LOGO_GTD) {
    logger.warn(`Conformity Sheet (offset ${rowOffset}): Logo path (RUTA_LOGO_GTD) is undefined. Skipping logo.`);
  } else {
    try {
      const logoNativeWidth = 64;
      const logoNativeHeight = 47;
      await insertResizedImage(
        sheet, RUTA_LOGO_GTD, actualLogoCellRef,
        logoNativeWidth, logoNativeHeight, true
      );
      logger.info(`Conformity Sheet (offset ${rowOffset}): Call to insertResizedImage completed for logo at ${actualLogoCellRef}.`);
    } catch (error) {
      logger.error(`Conformity Sheet (offset ${rowOffset}): Error calling insertResizedImage for GTD logo: ${error.message}`, { pathUsed: RUTA_LOGO_GTD, errorObj: error });
    }
  }

  // Set Acta Title
  const titleCellRef = offsetCellOrRangeRef(conformityConfig.celda_titulo_acta_base, rowOffset);
  applyCellStyles(sheet.getCell(titleCellRef), {
    value: conformitySheetTitle,
    font: COMMON_FONTS.title_arial_16_bold_center,
    alignment: COMMON_ALIGNMENTS.center_center_no_wrap
  });

  _applyFixedTextsConformity(sheet, rowOffset, conformityConfig, COMMON_FONTS, COMMON_ALIGNMENTS);

  // Populate general data
  Object.entries(MAIN_JSON_CELL_MAP_CONFORMIDAD).forEach(([jsonKey, baseCellRef]) => {
    const dataValue = generalData[jsonKey];
    if (dataValue !== undefined && dataValue !== null) {
      const actualCellRef = offsetCellOrRangeRef(baseCellRef, rowOffset);
      applyCellStyles(sheet.getCell(actualCellRef), {
        value: String(dataValue),
        font: COMMON_FONTS.data_arial_11_center,
        alignment: COMMON_ALIGNMENTS.center_center_wrap
      });
    }
  });

  _mergeCellsConformity(sheet, rowOffset, conformityConfig);

  // --- SECCIONES DE CONTENIDO PRINCIPAL (AHORA ACTIVAS) ---
  await _createPhotoCommentSections(sheet, rowOffset, photosForInstance, conformityConfig, { COMMON_FONTS, COMMON_ALIGNMENTS, BORDER_STYLES });
  _applySpecificBordersConformity(sheet, rowOffset, conformityConfig, BORDER_STYLES);
  
  const startOuterBorderRow = 5 + rowOffset;
  const endOuterBorderActualRow = conformityConfig.max_fila_contenido_bloque_base + rowOffset;
  applyOuterBorder(sheet, `B${startOuterBorderRow}:AG${endOuterBorderActualRow}`, BORDER_STYLES.THICK_SIDE);
  // --- FIN SECCIONES DE CONTENIDO PRINCIPAL ---

  // --- LÓGICA PARA ALTURAS DE FILA ESPECÍFICAS (111-114) CON NUEVA ALTURA ---
  const specificRowsToAdjust = [111, 112, 113, 114];
  const customHeight = 28; // <--- CAMBIADO A 28

  specificRowsToAdjust.forEach(baseRowNumber => {
    const actualRowNumber = baseRowNumber + rowOffset;
    sheet.getRow(actualRowNumber).height = customHeight;
    logger.debug(`Conformity Sheet (offset ${rowOffset}): Set row ${actualRowNumber} height to ${customHeight}`);
  });
  // --- FIN DE LA LÓGICA DE ALTURAS DE FILA ---

  logger.info(`Conformity sheet instance created with offset ${rowOffset}`);
}

export { createConformitySheetInstance };