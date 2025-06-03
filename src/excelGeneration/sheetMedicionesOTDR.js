// src/excelGeneration/sheetMedicionesOTDR.js

import config from '../config/index.js';
import {
  applyCellStyles,
  applyOuterBorder,
  applyFullBorderToRange,
  insertResizedImage
} from './excelUtils.js';
import logger from '../utils/logger.js'; // <--- Importamos el logger real aquí

// --- Constantes de Diseño y Datos Específicas para Mediciones OTDR ---
const PIXEL_TO_CHAR_FACTOR_COL_WIDTH = 7.0;

const otdrColWidths = {
    'A': 13 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH, 'B': 155 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH,
    'C': 84 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH, 'D': 84 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH,
    'E': 84 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH, 'F': 101 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH,
    'G': 101 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH, 'H': 90 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH,
    'I': 80 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH, 'J': 93 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH,
    'K': 13 / PIXEL_TO_CHAR_FACTOR_COL_WIDTH,
};

const otdrRowHeights = {
    1: 12.75, 2: 12.75, 3: 23.25, 4: 12.75, 5: 13.5, 6: 13.5, 7: 6.75, 8: 13.5, 9: 6.75, 10: 13.5,
    11: 6.75, 12: 13.5, 13: 6.75, 14: 13.5, 15: 6.75, 16: 13.5, 17: 6.75, 18: 13.5, 19: 13.5,
    20: 33.75
};

const dataInfoOTDR = [
    { labelText: "Cliente", labelCell: "B6", jsonKey: "Nombre cliente / Proyecto", dataAnchor: "C6", dataMergeRange: "C6:F6" },
    { labelText: "Distrito", labelCell: "G6", jsonKey: "Distrito", dataAnchor: "H6", dataMergeRange: "H6:J6" },
    { labelText: "N° Proyecto", labelCell: "B8", jsonKey: "N° PROY/ COD: AX", dataAnchor: "C8", dataMergeRange: "C8:F8" },
    { labelText: "Fecha", labelCell: "G8", jsonKey: null, dataAnchor: "H8", dataMergeRange: "H8:J8" },
    { labelText: "Nodo", labelCell: "B10", jsonKey: "Nodo", dataAnchor: "C10", dataMergeRange: "C10:D10" },
    { labelText: "Tipo de fibra", labelCell: "G10", jsonKey: null, dataAnchor: "H10", dataMergeRange: "H10:J10" },
    { labelText: "Medic. Desde", labelCell: "B12", jsonKey: null, dataAnchor: "C12", dataMergeRange: "C12:D12" },
    { labelText: "Hasta", labelCell: "G12", jsonKey: null, dataAnchor: "H12", dataMergeRange: "H12:I12" },
    { labelText: "Cable", labelCell: "B14", jsonKey: null, dataAnchor: "C14", dataMergeRange: "C14:E14", addYellowFill: true }, // <-- Marcado para fondo amarillo
    { labelText: "Ventana", labelCell: "G14", jsonKey: null, dataAnchor: "H14", dataMergeRange: "H14:I14", addYellowFill: true }  // <-- Marcado para fondo amarillo
];

const headersRow20 = {
    'B20': "Cuenta cable / Nodo / N° Cable", 'C20': "Nº Minitubo", 'D20': "Nº Fibra",
    'E20': "Cantidad Empalmes", 'F20': "Distancia (m)", 'G20': "Total enfrentadores",
    'H20': "Teórico Db / Km", 'I20': "Atenuacion Real", 'J20': "Cumple S/N"
};

/**
 * Creates and populates the "Mediciones OTDR" sheet.
 * @param {import('exceljs').Worksheet} sheet - The ExcelJS worksheet object.
 * @param {object} generalData - The general data loaded from datos.json.
 * @param {string} logoPath - Absolute path to the logo image.
 */
async function createSheetMedicionesOTDR(sheet, generalData, logoPath) {
  logger.info("Iniciando creación de hoja 'Mediciones OTDR'...");

  Object.entries(otdrColWidths).forEach(([col, width]) => {
    sheet.getColumn(col).width = width;
  });

  Object.entries(otdrRowHeights).forEach(([rowNum, height]) => {
    sheet.getRow(parseInt(rowNum, 10)).height = height;
  });

  if (logoPath) {
    try {
      const logoNativeWidth = 64;
      const logoNativeHeight = 47;
      await insertResizedImage(sheet, logoPath, 'B2', logoNativeWidth, logoNativeHeight, true);
      logger.info("Mediciones OTDR: Logo insertado en B2 via insertResizedImage.");
    } catch (e) {
      logger.error("Mediciones OTDR: Error al insertar logo.", e);
    }
  } else {
    logger.warn("Mediciones OTDR: No se proporcionó ruta para el logo.");
  }

  const { COMMON_FONTS, COMMON_ALIGNMENTS, BORDER_STYLES } = config.excel;
  const mediumOuterBorderSide = BORDER_STYLES.MEDIUM_SIDE || { style: 'medium' };
  const yellowFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; // Amarillo brillante

  sheet.mergeCells('B3:K3');
  applyCellStyles(sheet.getCell('B3'), {
    value: "MEDICIONES CABLE DE FIBRA OPTICA",
    font: COMMON_FONTS.title_arial_18_bold_center,
    alignment: COMMON_ALIGNMENTS.center_center_no_wrap
  });

  dataInfoOTDR.forEach(item => {
    applyCellStyles(sheet.getCell(item.labelCell), {
        value: item.labelText,
        font: COMMON_FONTS.otdr_label_10_bold,
        alignment: COMMON_ALIGNMENTS.center_center_wrap
    });

    const dataCell = sheet.getCell(item.dataAnchor); // Celda superior izquierda del merge de datos
    sheet.mergeCells(item.dataMergeRange);
    applyOuterBorder(sheet, item.dataMergeRange, mediumOuterBorderSide);

    if (item.jsonKey) {
        const dataValue = generalData[item.jsonKey] || "";
        applyCellStyles(dataCell, {
            value: String(dataValue),
            font: COMMON_FONTS.otdr_data_10_center,
            alignment: COMMON_ALIGNMENTS.center_center_wrap
        });
    }

    if (item.addYellowFill) { // Aplicar fondo si está marcado en la data
        dataCell.fill = yellowFill;
    }
  });

  const labelFontOtdr = COMMON_FONTS.otdr_label_10_bold;
  const centerAlignWrap = COMMON_ALIGNMENTS.center_center_wrap;

  // Fila 16: Atenuaciones
  sheet.mergeCells('B16:C16');
  applyCellStyles(sheet.getCell('B16'), { value: "Atenuac.x enfrentador (Db)", font: labelFontOtdr, alignment: centerAlignWrap });
  applyOuterBorder(sheet, 'D16:D16', mediumOuterBorderSide);
  sheet.getCell('D16').fill = yellowFill;

  sheet.mergeCells('E16:F16');
  applyCellStyles(sheet.getCell('E16'), { value: "Atenuación x empalme", font: labelFontOtdr, alignment: centerAlignWrap });
  applyOuterBorder(sheet, 'G16:G16', mediumOuterBorderSide);
  sheet.getCell('G16').fill = yellowFill;

  sheet.mergeCells('H16:I16');
  applyCellStyles(sheet.getCell('H16'), { value: "Atenuación Db / Km", font: labelFontOtdr, alignment: centerAlignWrap });
  applyOuterBorder(sheet, 'J16:J16', mediumOuterBorderSide);
  sheet.getCell('J16').fill = yellowFill;

  // Fila 18: Marca y Modelo OTDR
  applyCellStyles(sheet.getCell('B18'), { value: "Marca OTDR", font: labelFontOtdr, alignment: centerAlignWrap });
  sheet.mergeCells('C18:D18');
  applyOuterBorder(sheet, 'C18:D18', mediumOuterBorderSide);
  sheet.getCell('C18').fill = yellowFill;

  sheet.mergeCells('E18:F18');
  applyCellStyles(sheet.getCell('E18'), { value: "Modelo OTDR", font: labelFontOtdr, alignment: centerAlignWrap });
  sheet.mergeCells('G18:I18');
  applyOuterBorder(sheet, 'G18:I18', mediumOuterBorderSide);
  sheet.getCell('G18').fill = yellowFill;

  // Table Headers (Row 20)
  Object.entries(headersRow20).forEach(([cellRef, text]) => {
    applyCellStyles(sheet.getCell(cellRef), {
        value: text,
        font: COMMON_FONTS.otdr_label_10_bold,
        alignment: COMMON_ALIGNMENTS.center_center_wrap
    });
  });

  applyFullBorderToRange(sheet, 'B20:J164', BORDER_STYLES.BORDER_THIN_ALL_SIDES);
  applyOuterBorder(sheet, 'A1:K165', mediumOuterBorderSide);

  logger.info("Mediciones OTDR sheet created successfully.");
}

export { createSheetMedicionesOTDR };