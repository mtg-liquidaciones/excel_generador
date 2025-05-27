// src/excelGeneration/sheetActaResumen.js

import config from '../config/index.js';
import {
  applyCellStyles,
  applyOuterBorder,
  insertResizedImage
} from './excelUtils.js';

// Placeholder logger
const logger = {
  info: (message) => console.log(`[INFO] sheetActaResumen: ${message}`),
  warn: (message) => console.warn(`[WARN] sheetActaResumen: ${message}`),
  error: (message, errorDetails) => {
    const detailsString = errorDetails && typeof errorDetails === 'object' ? JSON.stringify(errorDetails) : errorDetails || '';
    console.error(`[ERROR] sheetActaResumen: ${message}`, detailsString);
  },
  debug: (message) => console.log(`[DEBUG] sheetActaResumen: ${message}`),
};

// --- Constantes de Diseño y Datos Específicas para Acta Resumen Pext ---
const PIXEL_TO_CHAR_FACTOR_COL_WIDTH = 7.0;

const actaResumenColWidths = {
    'A': 11 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH, 'B': 7 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH,
    'C': 21 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH, 'D': 133 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH,
    'E': 30 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH, 'F': 27 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH,
    'G': 20 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH, 'H': 43 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH,
    'I': 25 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH, 'J': 5 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH,
    'K': 25 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH, 'L': 6 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH,
    'M': 26 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH, 'N': 5 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH,
    'O': 39 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH, 'P': 13 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH,
    'Q': 156 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH, 'R': 103 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH,
    'S': 25 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH, 'T': 4 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH,
    'U': 25 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH, 'V': 5 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH,
    'W': 24 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH, 'X': 14 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH,
    'Y': 10 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH,
};

const actaResumenRowHeights = {
    1: 12.75, 2: 9.75, 3: 0.1, 4: 0.1, 5: 0.1, 6: 0.1, 7: 24, 8: 20.25, 9: 15.75, 10: 5.25,
    11: 15.75, 12: 4.5, 13: 15.75, 14: 3.75, 15: 15.73, 16: 3.75, 17: 15.73, 18: 3.75,
    19: 15.75, 20: 15.75, 21: 15.75, 22: 3.75, 23: 13.5, 24: 3.75, 25: 13.5, 26: 3.75, 27: 13.5, 28: 3.75, 29: 13.5, 30: 3.75,
    31: 13.5, 32: 3.75, 33: 15.75, 34: 3.75, 35: 13.5, 36: 3.75, 37: 13.5, 38: 3.75, 39: 13.5, 40: 3.75, 41: 13.5, 42: 3.75, 43: 13.5, 44: 3.75, 45: 13.5,
    46: 3.75, 47: 15.75, 48: 3.75, 49: 13.5, 50: 3.75, 51: 13.5, 52: 3.75, 53: 13.5, 54: 3.75, 55: 13.5, 56: 3.75, 57: 13.5, 58: 3.75, 59: 13.5, 60: 3.75,
    61: 13.5, 62: 3.75, 63: 6.75, 64: 22.5, 65: 13.5, 66: 4.5, 67: 12.5, 68: 12.5, 69: 3.75, 70: 13.5, 71: 3.75, 72: 13.5, 73: 3.75, 74: 18.75, 75: 18.75,
    76: 18.75, 77: 18.75, 78: 3.75, 79: 3.75, 80: 12.50, 81: 3.75, 82: 12.50, 83: 12.50, 84: 3.75
};

const dataForSummaryWithJsonKeys = [
    { labelText: "Ciudad", labelCell: "Q9", jsonKey: "Ciudad", dataRange: "R9:W9" },
    { labelText: "Contratista", labelCell: "D11", jsonKey: "Contratista", dataRange: "E11:O11" },
    { labelText: "N° PROY/ COD: AX", labelCell: "Q11", jsonKey: "N° PROY/ COD: AX", dataRange: "R11:W11" },
    { labelText: "Distrito", labelCell: "D13", jsonKey: "Distrito", dataRange: "E13:O13" },
    { labelText: "Fecha inicio", labelCell: "Q13", jsonKey: "Fecha Inicio", dataRange: "R13:W13" },
    { labelText: "Nodo", labelCell: "D15", jsonKey: "Nodo", dataRange: "E15:O15" },
    { labelText: "Fecha termino", labelCell: "Q15", jsonKey: "Fecha Término", dataRange: "R15:W15" },
    { labelText: "Nombre cliente / Proyecto", labelCell: "D17", jsonKey: "Nombre cliente / Proyecto", dataRange: "H17:W17", preMergeLabel: "D17:G17" },
    { labelText: "Direccion Cliente", labelCell: "D19", jsonKey: "Direccion Cliente", dataRange: "E19:W19" }
];

const checklistLayout = [
    { row: 23, dText: "Tendidos correctamente ordenado", oText: "Mufa corresponde a la suminstrada", dHasCheck: true, oHasCheck: true },
    { row: 25, dText: "Placas de identificación en cámaras (pdte)", oText: "Minitubos ordenados", dHasCheck: true, oHasCheck: true },
    { row: 27, dText: "Cable amarrado a soportes en camara", oText: "Fusiones correctamente realizadas y ordenadas", dHasCheck: true, oHasCheck: true },
    { row: 29, dText: "Reservas ( 2 vueltas en interior de camara)", oText: "Uso correcto de código colores", dHasCheck: true, oHasCheck: true },
    { row: 31, dText: "Limpieza de camara ( residuos)", oText: "Terminación correcta del Keplar", dHasCheck: true, oHasCheck: true },
    
    { row: 33, dText: "Tendido Aereo", dFontKey: "fontArial11Bold", dMerge: 'D33:H33', dHasCheck: false, oText: "Pruebas de FO y archivo en digital", oFontKey: "fontArial10", oMerge: 'O33:R33', oHasCheck: true },
    { row: 35, dText: "Tendidos correctamente ordenado", oText: "Identificacion cable FO Origen-Destino en Mufa FO", dHasCheck: true, oHasCheck: true },
    { row: 37, dText: "Placas de identificación (pdte)", oText: "Otros Trabajos", oFontKey: "fontArial11Bold", dHasCheck: true, oHasCheck: false }, // Corregido oHasCheck a false según tu indicación
    { row: 39, dText: "Correcta terminacion de preformada", oText: "Colocación correcta de anclas", dHasCheck: true, oHasCheck: true },
    { row: 41, dText: "Correcta terminacion de suspensión", oText: "Instalación de mensajero adecuado", dHasCheck: true, oHasCheck: true },
    { row: 43, dText: "Reservas", oText: "Altura cruce de calles correcto", dHasCheck: true, oHasCheck: true },
    { row: 45, dText: "Limpieza de zona de trabajo", oText: "Colocacion Flejes en Bajadas de poste ", dHasCheck: true, oHasCheck: true },

    { row: 47, dText: "Cabeceras y Rack", dFontKey: "fontArial11Bold", dMerge: 'D47:H47', dHasCheck: false, oText: "Sellado ductos en Bajada de poste", oFontKey: "fontArial10", oMerge: 'O47:R47', oHasCheck: true },
    { row: 49, dText: "ODF Nodo correctamente terminadas", oText: "Postes alineados y en buen estado", dHasCheck: true, oHasCheck: true },
    { row: 51, dText: "ODF Cliente perfectamente instalada", oText: "Instalación de cajas verticales para Edificios", dHasCheck: true, oHasCheck: true },
    { row: 53, dText: "Identificación ODF", oText: "Entrega Asbuilt", oFontKey: "fontArial11Bold", dHasCheck: true, oHasCheck: true }, // *** CORREGIDO: oHasCheck: true ***
    { row: 55, dText: "Rack instalados", oText: "Diseño", oFontKey: "fontArial11Bold", dHasCheck: true, oHasCheck: false }, // Corregido oHasCheck a false según tu indicación
    { row: 57, dText: "Ordenadores horizontales", oText: "Diseño de acuerdo a lo Solicitado", dHasCheck: true, oHasCheck: true },
    { row: 59, dText: "Ordenadores verticales", oText: "Cumple con la Norma de Dibujo GTD Wigo", dHasCheck: true, oHasCheck: true },
    { row: 61, dText: "", dHasCheck: false, oText: "Entrega Cartas Ingresos en entidades del Estado", oFontKey: "fontArial10", oMerge: 'O61:R61', oHasCheck: true } 
  ];

/**
 * Creates and populates the "Acta Resumen Pext" sheet.
 * @param {import('exceljs').Worksheet} sheet - The ExcelJS worksheet object.
 * @param {object} generalData - The general data loaded from datos.json.
 * @param {string} logoPath - Absolute path to the logo image.
 */
async function createSheetActaResumen(sheet, generalData, logoPath) {
  logger.info("Iniciando creación de hoja 'Acta Resumen Pext' (versión detallada)...");

  const { COMMON_FONTS, COMMON_ALIGNMENTS, BORDER_STYLES } = config.excel;

  const definedFonts = {
    fontArial16Bold: { name: 'Arial', size: 16, bold: true },
    fontArial11Bold: { name: 'Arial', size: 11, bold: true },
    fontArial10Bold: { name: 'Arial', size: 10, bold: true },
    fontArial10: { name: 'Arial', size: 10 },
    fontArial8: { name: 'Arial', size: 8 },
    fontDataDefault: COMMON_FONTS.data_arial_11_center || { name: 'Arial', size: 11 }
  };


  logger.debug("Aplicando anchos de columna...");
  Object.entries(actaResumenColWidths).forEach(([col, calculatedWidth]) => {
    sheet.getColumn(col).width = calculatedWidth;
  });

  Object.entries(actaResumenRowHeights).forEach(([rowNum, height]) => {
    sheet.getRow(parseInt(rowNum, 10)).height = height;
  });

  if (logoPath) {
    try {
      const logoNativeWidth = 64;
      const logoNativeHeight = 47;
      await insertResizedImage(sheet, logoPath, 'C7', logoNativeWidth, logoNativeHeight, true);
      logger.info("Acta Resumen: Logo insertado en C7 via insertResizedImage.");
    } catch (e) {
      logger.error("Acta Resumen: Error al insertar logo via insertResizedImage.", e);
    }
  } else {
    logger.warn("Acta Resumen: No se proporcionó ruta para el logo.");
  }

  sheet.mergeCells('D7:X7');
  applyCellStyles(sheet.getCell('D7'), { value: "ACTA DE ENTREGA DE OBRA", font: definedFonts.fontArial16Bold, alignment: COMMON_ALIGNMENTS.center_center_no_wrap });
  sheet.mergeCells('F8:Q8');
  applyCellStyles(sheet.getCell('F8'), { value: "TENDIDO", font: definedFonts.fontArial16Bold, alignment: COMMON_ALIGNMENTS.center_center_no_wrap });

  const mediumOuterBorderSide = BORDER_STYLES.MEDIUM_SIDE || { style: 'medium' };
  const mediumBottomBorder = { bottom: mediumOuterBorderSide };

  dataForSummaryWithJsonKeys.forEach(item => {
    try {
      applyCellStyles(sheet.getCell(item.labelCell), { value: item.labelText, font: definedFonts.fontArial10Bold, alignment: COMMON_ALIGNMENTS.center_center_wrap });
      if (item.preMergeLabel) {
        if (typeof item.preMergeLabel === 'string' && item.preMergeLabel.includes(':')) {
          sheet.mergeCells(item.preMergeLabel);
        } else { logger.warn(`Valor de preMergeLabel inválido para ${item.labelText}: ${item.preMergeLabel}`); }
      }
      if (typeof item.dataRange === 'string' && item.dataRange.includes(':')) {
        sheet.mergeCells(item.dataRange);
        applyOuterBorder(sheet, item.dataRange, mediumOuterBorderSide);
        const dataValue = generalData[item.jsonKey] || "";
        applyCellStyles(sheet.getCell(item.dataRange.split(':')[0]), { value: String(dataValue), font: definedFonts.fontDataDefault, alignment: COMMON_ALIGNMENTS.center_center_wrap });
      } else { logger.warn(`Valor de dataRange inválido para ${item.labelText}: ${item.dataRange}`); }
    } catch (e) { logger.error(`Error procesando item de cabecera '${item.labelText}': ${e.message}`, item); }
  });

  logger.debug("Construyendo sección de Checklist...");

  applyCellStyles(sheet.getCell('D21'), { value: "Tendido subterráneo Cables", font: definedFonts.fontArial11Bold }); sheet.mergeCells('D21:H21');
  applyCellStyles(sheet.getCell('I21'), { value: "Si", font: definedFonts.fontArial10Bold, alignment: COMMON_ALIGNMENTS.center_center_wrap });
  applyCellStyles(sheet.getCell('K21'), { value: "No", font: definedFonts.fontArial10Bold, alignment: COMMON_ALIGNMENTS.center_center_wrap });
  applyCellStyles(sheet.getCell('M21'), { value: "NA", font: definedFonts.fontArial10Bold, alignment: COMMON_ALIGNMENTS.center_center_wrap });
  applyCellStyles(sheet.getCell('O21'), { value: "Empalmes FO", font: definedFonts.fontArial11Bold }); sheet.mergeCells('O21:R21');
  applyCellStyles(sheet.getCell('S21'), { value: "Si", font: definedFonts.fontArial10Bold, alignment: COMMON_ALIGNMENTS.center_center_wrap });
  applyCellStyles(sheet.getCell('U21'), { value: "No", font: definedFonts.fontArial10Bold, alignment: COMMON_ALIGNMENTS.center_center_wrap });
  applyCellStyles(sheet.getCell('W21'), { value: "NA", font: definedFonts.fontArial10Bold, alignment: COMMON_ALIGNMENTS.center_center_wrap });

  const checkboxBorder = { 
    top: mediumOuterBorderSide, left: mediumOuterBorderSide, bottom: mediumOuterBorderSide, right: mediumOuterBorderSide 
  };

  checklistLayout.forEach(item => {
    const dColMerge = item.dMerge || `D${item.row}:H${item.row}`;
    const oColMerge = item.oMerge || `O${item.row}:R${item.row}`;
    const dFontToUse = item.dFontKey ? definedFonts[item.dFontKey] : definedFonts.fontArial10;
    const oFontToUse = item.oFontKey ? definedFonts[item.oFontKey] : definedFonts.fontArial10;

    if (item.dText) {
      applyCellStyles(sheet.getCell(`D${item.row}`), { value: item.dText, font: dFontToUse });
      sheet.mergeCells(dColMerge);
    }

    if (item.oText) {
      applyCellStyles(sheet.getCell(`O${item.row}`), { value: item.oText, font: oFontToUse });
      sheet.mergeCells(oColMerge);
    }

    if (item.dHasCheck) {
      ['I', 'K', 'M'].forEach(col => {
        const cell = sheet.getCell(`${col}${item.row}`);
        if (cell) cell.border = checkboxBorder;
      });
    } else {
      ['I', 'K', 'M'].forEach(col => {
        const cell = sheet.getCell(`${col}${item.row}`);
        if (cell) cell.border = undefined;
      });
    }

    if (item.oHasCheck) {
      ['S', 'U', 'W'].forEach(col => {
        const cell = sheet.getCell(`${col}${item.row}`);
        if (cell) cell.border = checkboxBorder;
      });
    } else {
      ['S', 'U', 'W'].forEach(col => {
        const cell = sheet.getCell(`${col}${item.row}`);
        if (cell) cell.border = undefined;
      });
    }
  });
  
  applyCellStyles(sheet.getCell('D65'), { value: "Nombre Supervisor Contratista", font: definedFonts.fontArial11Bold });
  sheet.mergeCells('H65:P65'); sheet.getCell('H65').border = mediumBottomBorder; 
  applyCellStyles(sheet.getCell('Q65'), { value: "Firma", font: definedFonts.fontArial11Bold, alignment: COMMON_ALIGNMENTS.right_center_no_wrap });
  sheet.mergeCells('R65:U65'); sheet.getCell('R65').border = mediumBottomBorder;

  applyCellStyles(sheet.getCell('D70'), { value: "Nombre Supervisor", font: definedFonts.fontArial11Bold });
  sheet.mergeCells('H70:P70'); sheet.getCell('H70').border = mediumBottomBorder;
  applyCellStyles(sheet.getCell('Q70'), { value: "Firma", font: definedFonts.fontArial11Bold, alignment: COMMON_ALIGNMENTS.right_center_no_wrap });
  sheet.mergeCells('R70:U70'); sheet.getCell('R70').border = mediumBottomBorder;

  applyCellStyles(sheet.getCell('D72'), { value: "OBSERVACIONES SUPERVISOR", font: definedFonts.fontArial11Bold }); sheet.mergeCells('D72:H72');
  applyCellStyles(sheet.getCell('Q72'), { value: "OBSERVACIONES CONTRALOR", font: definedFonts.fontArial11Bold }); sheet.mergeCells('Q72:R72');

  const dottedBottomBorder = { bottom: BORDER_STYLES.DOTTED_SIDE };
  ['D74:O74', 'P74:W74', 'D75:O75', 'P75:W75', 'D76:O76', 'P76:W76'].forEach(range => {
    if (typeof range === 'string' && range.includes(':')) {
      sheet.mergeCells(range);
      sheet.getCell(range.split(':')[0]).border = dottedBottomBorder;
    } else { logger.warn(`Rango inválido para borde punteado: ${range}`); }
  });
  sheet.mergeCells('D77:O77'); sheet.getCell('D77').border = { ...dottedBottomBorder };
  sheet.mergeCells('P77:W77'); sheet.getCell('P77').border = { ...dottedBottomBorder };

  applyOuterBorder(sheet, 'D74:O77', mediumOuterBorderSide);
  applyOuterBorder(sheet, 'P74:W77', mediumOuterBorderSide);

  applyCellStyles(sheet.getCell('C80'), { value: "* Deben ser completados todos los campos de observacion, siendo responsabilidad del Supervisor Despliegue y Personal de SCM", font: definedFonts.fontArial10 });
  sheet.mergeCells('C80:X80');
  applyCellStyles(sheet.getCell('C82'), { value: "** Los campos SI (Aceptado), NO (Rechazado) y NA (No aplica) deben ser marcados con una \"X\".", font: definedFonts.fontArial10 });
  sheet.mergeCells('C82:X82');
  applyCellStyles(sheet.getCell('C83'), { value: "Versión 2016-09", font: definedFonts.fontArial8 });
  sheet.mergeCells('C83:D83');

  applyOuterBorder(sheet, 'B2:X84', mediumOuterBorderSide);

  logger.info("Acta Resumen Pext sheet creation logic (detailed spec) fully completed.");
}

export { createSheetActaResumen };