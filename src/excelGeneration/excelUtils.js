// src/excelGeneration/excelUtils.js

import sharp from 'sharp'; // Para redimensionar imágenes
import fs from 'fs/promises';
import path from 'path'; // Útil para extensiones de archivo, etc.
import logger from '../utils/logger.js'; // <--- Importamos el logger real aquí
import config from '../config/index.js'; // Importamos la configuración para acceder a PHOTO_CONFIG

/**
 * Convierte un número de columna (1-indexed) a su letra correspondiente (A, B, ..., Z, AA, etc.).
 * @param {number} colNumber - El número de la columna (ej. 1 para A).
 * @returns {string} La letra de la columna.
 */
function getColumnLetter(colNumber) {
  let letter = '';
  let num = colNumber;
  while (num > 0) {
    const remainder = (num - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter; // 65 es el código ASCII de 'A'
    num = Math.floor((num - 1) / 26);
  }
  return letter;
}

/**
 * Convierte una letra de columna (A, B, ..., Z, AA, etc.) a su número correspondiente (1-indexed).
 * @param {string} colLetter - La letra de la columna (ej. "A").
 * @returns {number} El número de la columna.
 */
function getColumnNumber(colLetter) {
  let number = 0;
  const upperColLetter = colLetter.toUpperCase();
  for (let i = 0; i < upperColLetter.length; i++) {
    number = number * 26 + (upperColLetter.charCodeAt(i) - 64); // 64 = 'A'.charCodeAt(0) - 1
  }
  return number;
}

/**
 * Parsea una referencia de celda (ej. "A1") en sus componentes de columna y fila.
 * @param {string} cellRef - La referencia de la celda.
 * @returns {{ col: string, row: number, colNum: number }}
 * @throws Error si la referencia es inválida.
 */
function parseCellRef(cellRef) {
  const match = cellRef.match(/([A-Z]+)([0-9]+)/i);
  if (!match) {
    throw new Error(`Referencia de celda inválida: ${cellRef}`);
  }
  const col = match[1].toUpperCase();
  const row = parseInt(match[2], 10);
  return { col, row, colNum: getColumnNumber(col) };
}

/**
 * Parsea un string de rango (ej. "A1:B5" o "A1") en sus celdas de inicio y fin.
 * @param {string} rangeString - El string del rango.
 * @returns {{ start: { col: string, row: number, colNum: number }, end: { col: string, row: number, colNum: number } }}
 * @throws Error si el rango es inválido.
 */
function parseRangeString(rangeString) {
  if (!rangeString.includes(':')) {
    const parsedCell = parseCellRef(rangeString);
    return { start: parsedCell, end: parsedCell };
  }
  const parts = rangeString.split(':');
  if (parts.length !== 2) {
    throw new Error(`String de rango inválido: ${rangeString}`);
  }
  return { start: parseCellRef(parts[0]), end: parseCellRef(parts[1]) };
}

/**
 * Aplica un desplazamiento de fila y/o columna a una referencia de celda o rango.
 * @param {string} originalRef - La referencia original (ej. "A1" o "A1:B5").
 * @param {number} rowOffset - El desplazamiento a aplicar a las filas.
 * @param {number} [colOffset=0] - El desplazamiento a aplicar a las columnas.
 * @returns {string} La nueva referencia de celda o rango.
 */
function offsetCellOrRangeRef(originalRef, rowOffset, colOffset = 0) {
  const offsetSingle = (ref) => {
    const { colNum, row } = parseCellRef(ref);
    const newRowNum = row + rowOffset;
    const newColNum = colNum + colOffset;
    if (newRowNum < 1 || newColNum < 1) {
      logger.warn(`Offset para '${ref}' con rowOffset ${rowOffset}, colOffset ${colOffset} resultó en coordenadas inválidas: Fila ${newRowNum}, Col ${newColNum}. Se intentará ajustar a 1.`);
      return `${getColumnLetter(Math.max(1, newColNum))}${Math.max(1, newRowNum)}`;
    }
    return `${getColumnLetter(newColNum)}${newRowNum}`;
  };

  if (originalRef.includes(':')) {
    const [startCell, endCell] = originalRef.split(':');
    return `${offsetSingle(startCell)}:${offsetSingle(endCell)}`;
  } else {
    return offsetSingle(originalRef);
  }
}

/**
 * Aplica estilos a una celda.
 * @param {import('exceljs').Cell} cell - El objeto celda de ExcelJS.
 * @param {object} styles - Objeto con los estilos a aplicar.
 * @param {any} [styles.value] - El valor de la celda.
 * @param {Partial<import('exceljs').Font>} [styles.font] - Estilo de fuente.
 * @param {Partial<import('exceljs').Alignment>} [styles.alignment] - Estilo de alineación.
 * @param {Partial<import('exceljs').Borders>} [styles.border] - Estilo de borde completo.
 * @param {import('exceljs').Fill} [styles.fill] - Estilo de relleno.
 * @param {string} [styles.numFmt] - Formato numérico (ej. '0.00%').
 */
function applyCellStyles(cell, styles = {}) {
  if (styles.value !== undefined) cell.value = styles.value;
  if (styles.font) cell.font = styles.font;
  if (styles.alignment) cell.alignment = styles.alignment;
  if (styles.border) cell.border = styles.border;
  if (styles.fill) cell.fill = styles.fill;
  if (styles.numFmt) cell.numFmt = styles.numFmt;
}

/**
 * Inserta una imagen redimensionada en la hoja.
 * @param {import('exceljs').Worksheet} sheet - La hoja de ExcelJS.
 * @param {string} imagePath - Ruta absoluta a la imagen.
 * @param {string} cellAnchor - Celda donde anclar la esquina superior izquierda de la imagen (ej. "A1").
 * @param {number} targetPixelWidth - Ancho deseado en píxeles.
 * @param {number} targetPixelHeight - Alto deseado en píxeles.
 * @param {boolean} [maintainAspectRatio=false] - Si se debe mantener la relación de aspecto.
 */
async function insertResizedImage(sheet, imagePath, cellAnchor, targetPixelWidth, targetPixelHeight, maintainAspectRatio = false) {
  try {
    await fs.access(imagePath);
    const imageBuffer = await fs.readFile(imagePath);
    const image = sharp(imageBuffer);
    const metadata = await image.metadata();

    let newWidth = Math.max(1, Math.round(targetPixelWidth));
    let newHeight = Math.max(1, Math.round(targetPixelHeight));

    // Si maintainAspectRatio es true (ej. para logos), calcula nuevas dimensiones manteniendo el aspecto.
    // Si es false (como para las fotos de conformidad), usa las dimensiones targetPixelWidth y targetPixelHeight directamente
    // y fuerza el ajuste 'fill' para que la imagen complete el espacio sin recortar.
    let sharpResizeOptions = {
        width: newWidth,
        height: newHeight,
    };

    if (maintainAspectRatio) {
        // Para el logo, queremos que se ajuste sin distorsionarse, manteniendo el aspecto.
        // El 'contain' intentará encajar la imagen dentro de las dimensiones dadas.
        sharpResizeOptions.fit = sharp.fit.contain;
        // Si el objetivo es un ajuste con aspecto para un logo, podemos permitir que sharp
        // recalcule el ancho/alto final para asegurar el aspecto dentro de las dimensiones.
        // Aquí no estamos usando metadata.width/height para el cálculo inicial si maintainAspectRatio es true
        // sino para la proporción. sharp.fit.contain lo maneja.
        logger.debug(`Redimensionando imagen con aspecto (fit: contain): Original ${metadata.width}x${metadata.height}, Target ${targetPixelWidth}x${targetPixelHeight}, Nuevo ${newWidth}x${newHeight}`);
    } else {
        // Para las fotos de conformidad, queremos que la imagen llene exactamente el espacio,
        // distorsionándose si es necesario para ajustarse a las dimensiones exactas.
        sharpResizeOptions.fit = sharp.fit.fill;
        logger.debug(`Redimensionando imagen sin mantener aspecto (fit: fill): Target ${newWidth}x${newHeight}`);
    }

    const outputFormat = (metadata.format === 'jpeg' || metadata.format === 'jpg') ? 'jpeg' : 'png';
    const resizedImageBuffer = await image.resize(sharpResizeOptions)[outputFormat]().toBuffer();

    const imageId = sheet.workbook.addImage({
      buffer: resizedImageBuffer,
      extension: outputFormat,
    });

    const anchorPos = parseCellRef(cellAnchor); // .colNum y .row son 1-indexados
    sheet.addImage(imageId, {
      // tl (top-left) usa coordenadas 0-indexadas para la celda
      tl: { col: anchorPos.colNum - 1, row: anchorPos.row - 1 },
      ext: { width: newWidth, height: newHeight },
    });
    logger.debug(`Imagen ${imagePath} insertada en ${cellAnchor} con dimensiones finales ${newWidth}x${newHeight}`);
  } catch (error) {
    logger.error(`Error al procesar/insertar imagen '${imagePath}' en '${cellAnchor}': ${error.message}`, error);
  }
}

/**
 * Aplica un estilo de borde a los cuatro lados exteriores de un rango de celdas.
 * Preserva los estilos de borde existentes en los otros lados de las celdas del perímetro.
 * @param {import('exceljs').Worksheet} sheet - La hoja de ExcelJS.
 * @param {string} rangeString - El rango (ej. "A1:C5").
 * @param {Partial<import('exceljs').BorderSide>} borderSideStyle - El estilo de borde a aplicar (ej. { style: 'thin' }).
 */
function applyOuterBorder(sheet, rangeString, borderSideStyle) {
  try {
    const { start, end } = parseRangeString(rangeString);
    if (!borderSideStyle || Object.keys(borderSideStyle).length === 0) {
      logger.warn(`applyOuterBorder llamado para ${rangeString} sin borderSideStyle válido.`);
      return;
    }

    for (let r = start.row; r <= end.row; r++) {
      for (let c = start.colNum; c <= end.colNum; c++) {
        if (r === start.row || r === end.row || c === start.colNum || c === end.colNum) {
          const cell = sheet.getCell(r, c);
          const currentBorder = cell.border || {}; // Obtener borde actual o un objeto vacío
          const newBorder = { ...currentBorder };

          if (r === start.row) newBorder.top = { ...currentBorder.top, ...borderSideStyle };
          if (r === end.row) newBorder.bottom = { ...currentBorder.bottom, ...borderSideStyle };
          if (c === start.colNum) newBorder.left = { ...currentBorder.left, ...borderSideStyle };
          if (c === end.colNum) newBorder.right = { ...currentBorder.right, ...borderSideStyle };

          // Asegurar que no se apliquen lados vacíos si no estaban definidos y borderSideStyle no los cubre
          Object.keys(newBorder).forEach(key => {
            if (newBorder[key] && Object.keys(newBorder[key]).length === 0) {
              delete newBorder[key]; // Evita { style: undefined } etc.
            }
          });

          if(Object.keys(newBorder).length > 0) {
            cell.border = newBorder;
          }
        }
      }
    }
  } catch (e) {
    logger.error(`Error en applyOuterBorder para rango ${rangeString}: ${e.message}`);
  }
}

/**
 * Aplica un objeto de borde completo a cada celda dentro de un rango.
 * @param {import('exceljs').Worksheet} sheet - La hoja de ExcelJS.
 * @param {string} rangeString - El rango (ej. "A1:C5").
 * @param {Partial<import('exceljs').Borders>} borderStyle - El objeto de borde completo (ej. { top: {...}, left: {...}, ...}).
 */
function applyFullBorderToRange(sheet, rangeString, borderStyle) {
  try {
    const { start, end } = parseRangeString(rangeString);
    if (!borderStyle || Object.keys(borderStyle).length === 0) {
      logger.warn(`applyFullBorderToRange llamado para ${rangeString} sin borderStyle válido.`);
      return;
    }
    for (let r = start.row; r <= end.row; r++) {
      for (let c = start.colNum; c <= end.colNum; c++) {
        sheet.getCell(r, c).border = borderStyle;
      }
    }
  } catch (e) {
    logger.error(`Error en applyFullBorderToRange para rango ${rangeString}: ${e.message}`);
  }
}

export {
  getColumnLetter,
  getColumnNumber,
  parseCellRef,
  parseRangeString,
  offsetCellOrRangeRef,
  applyCellStyles,
  insertResizedImage,
  applyOuterBorder,
  applyFullBorderToRange,
};