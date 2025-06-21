// src/config/styleAndLayoutConfig.js

// NOTA: Las rutas absolutas como RUTA_LOGO_GTD podrían ser mejor manejadas
// a través de variables de entorno o haciéndolas relativas al proyecto.
const RUTA_LOGO_GTD = process.env.RUTA_LOGO_GTD || "/var/www/excel_generador/logogtd.png";

const NOMBRE_ARCHIVO_DATOS_PRINCIPAL = "datos.json";
const NOMBRE_ARCHIVO_COMENTARIOS = "comentarios.json";

const FOLDER_TO_TITLE_MAP = {
  "Entrega de Nodos": "ACTA CONFORMIDAD ENTREGA DE NODOS",
  "Cables Aereo y Subterraneos": "ACTA CONFORMIDAD TENDIDO AEREO Y SUBTERRANEO",
  "Reservas Aereas": "ACTA CONFORMIDAD MUFAS, RESERVAS AEREAS Y CRUCES DE CALLE O AVENIDA",
  "OO.CC - CAMARAS": "ACTA CONFORMIDAD OBRA CIVIL - CAMARAS",
  "OO.CC - POSTES": "ACTA CONFORMIDAD OBRA CIVIL - INST.POSTES",
  "OO.CC - CANALIZADO": "ACTA CONFORMIDAD OBRA CIVIL - CANALIZACION",
  "Empalmes": "ACTA CONFORMIDAD FUSIONES",
  "Manipulación de MUFA": "ACTA CONFORMIDAD MANIPULACIÓN DE MUFA"
};

const MAIN_JSON_CELL_MAP_CONFORMIDAD = {
  "Contratista": "E10", "Distrito": "E12", "Nodo": "E14",
  "Nombre cliente / Proyecto": "H16", "Direccion Cliente": "E18",
  "Ciudad": "T8", "N° PROY/ COD: AX": "T10",
  "Fecha Inicio": "T12", "Fecha Término": "T14"
};

// --- CONFIGURACIÓN DE FOTOS ---
const PHOTO_CONFIG = {
  TARGET_HEIGHT_CM: 9.9, // Altura deseada en CM
  TARGET_WIDTH_CM: 11.85, // Ancho deseado en CM
  // Factor de conversión de CM a Píxeles (asumiendo 96 DPI, estándar en muchas pantallas)
  // 1 pulgada = 2.54 cm. Si 96 DPI (píxeles por pulgada), entonces (96 / 2.54) píxeles por cm.
  CM_TO_PIXELS_FACTOR: (96 / 2.54),
  // Dimensiones en píxeles calculadas a partir de los CM exactos
  PHOTO_PIXEL_HEIGHT: Math.round(9.9 * (96 / 2.54)), // Altura en píxeles
  PHOTO_PIXEL_WIDTH: Math.round(11.85 * (96 / 2.54)), // Ancho en píxeles
  PHOTOS_PER_INSTANCE_STRUCTURE: 6,
  POSSIBLE_IMAGE_EXTENSIONS: ['.jpg', '.jpeg', '.png', '.gif', '.bmp']
};

// --- ESTILOS COMUNES (para ExcelJS) ---
// Documentación de estilos de ExcelJS:
// Font: { name: 'Calibri', family: 2, size: 11, bold: false, italic: false, underline: false, strike: false, color: { argb: 'FF000000' }, scheme: 'minor' }
// Alignment: { horizontal: 'left'|'center'|'right', vertical: 'top'|'middle'|'bottom', wrapText: false, shrinkToFit: false, indent: 0, textRotation: 0, readingOrder: 'ltr' }
// Border: { style: 'thin'|'medium'|'thick'|'dotted'|..., color: { argb: 'FF000000' } }
// Full Border Object: { top: BORDER_SIDE_OBJECT, left: BORDER_SIDE_OBJECT, bottom: BORDER_SIDE_OBJECT, right: BORDER_SIDE_OBJECT }

const COMMON_FONTS_EXCELJS = {
  label_bold_10: { name: 'Arial', size: 10, bold: true },
  data_arial_11_center: { name: 'Arial', size: 11 }, // La alineación se aplica por separado
  comment_arial_11_center: { name: 'Arial', size: 11 }, // La alineación se aplica por separado
  title_arial_16_bold_center: { name: 'Arial', size: 16, bold: true },
  title_arial_18_bold_center: { name: 'Arial', size: 18, bold: true },
  footnote_10: { name: 'Arial', size: 10 },
  version_8: { name: 'Arial', size: 8 },
  otdr_label_10_bold: { name: 'Arial', size: 10, bold: true },
  otdr_data_10_center: { name: 'Arial', size: 10 },
  otdr_header_11_bold: { name: 'Arial', size: 11, bold: true },
};

const COMMON_ALIGNMENTS_EXCELJS = {
  center_center_wrap: { horizontal: 'center', vertical: 'middle', wrapText: true },
  center_center_no_wrap: { horizontal: 'center', vertical: 'middle', wrapText: false },
  left_center_wrap: { horizontal: 'left', vertical: 'middle', wrapText: true },
  right_center_no_wrap: { horizontal: 'right', vertical: 'middle', wrapText: false },
  align_right: { horizontal: 'right', vertical: 'middle' }, // Añadido para S99, S102
};

const BORDER_STYLES_EXCELJS = {
  THIN_SIDE: { style: 'thin', color: { argb: 'FF000000' } }, // Negro por defecto
  MEDIUM_SIDE: { style: 'medium', color: { argb: 'FF000000' } },
  THICK_SIDE: { style: 'thick', color: { argb: 'FF000000' } },
  DOTTED_SIDE: { style: 'dotted', color: { argb: 'FF000000' } },

  BORDER_THIN_ALL_SIDES: {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } }
  },
  BORDER_MEDIUM_ALL_SIDES: {
    top: { style: 'medium', color: { argb: 'FF000000' } },
    left: { style: 'medium', color: { argb: 'FF000000' } },
    bottom: { style: 'medium', color: { argb: 'FF000000' } },
    right: { style: 'medium', color: { argb: 'FF000000' } }
  }
};

// --- CONFIGURACIÓN BASE PARA HOJAS DE ACTAS DE CONFORMIDAD ---
const _anchos_columnas_conformidad_px = {
  'A': 11, 'B': 7, 'C': 21, 'D': 149, 'E': 30, 'F': 36, 'G': 20, 'H': 43, 'I': 20, 'J': 5,
  'K': 5, 'L': 25, 'M': 6, 'N': 26, 'O': 5, 'P': 77, 'Q': 40, 'R': 13, 'S': 116, 'T': 103,
  'U': 25, 'V': 4, 'W': 25, 'X': 5, 'Y': 24, 'Z': 24, 'AA': 24, 'AB': 24, 'AC': 24, 'AD': 24,
  'AE': 24, 'AF': 24, 'AG': 14, 'AH': 10
};

// Convertir anchos de píxeles a "unidades de caracteres" para ExcelJS
// El factor 7.0 es una aproximación tomada del código Python.
const _PIXEL_TO_CHAR_FACTOR_COL_WIDTH = 7.0;
const anchos_columnas_conformidad_char = {};
for (const col in _anchos_columnas_conformidad_px) {
  anchos_columnas_conformidad_char[col] = _anchos_columnas_conformidad_px[col] / _PIXEL_TO_CHAR_FACTOR_COL_WIDTH;
}

const CONFORMIDAD_SHEET_CONFIG = {
  celda_logo_gtd: "C6",
  anchos_columnas_char: anchos_columnas_conformidad_char,
  alturas_filas_base: { // En puntos, igual que openpyxl
    1: 13, 2: 0.1, 3: 0.1, 4: 0.1, 5: 4.5, 6: 24, 7: 20.25, 8: 15.75, 9: 4.5, 10: 15.75,
    11: 4.5, 12: 15.75, 13: 4.5, 14: 15.75, 15: 4.5, 16: 15.75, 17: 4.5, 18: 15.75,
    19: 15.75, 20: 15.75
  },
  textos_fijos_base: { // Las claves 'font_key' y 'alignment_key' se usarán para buscar en COMMON_FONTS_EXCELJS y COMMON_ALIGNMENTS_EXCELJS
    'D10': { texto: 'Contratista', font_key: 'label_bold_10' },
    'D12': { texto: 'Distrito', font_key: 'label_bold_10' },
    'D14': { texto: 'Nodo', font_key: 'label_bold_10' },
    'D16': { texto: 'Nombre cliente / Proyecto', font_key: 'label_bold_10' },
    'D18': { texto: 'Direccion Cliente', font_key: 'label_bold_10' },
    'S8': { texto: 'Ciudad', font_key: 'label_bold_10' },
    'S10': { texto: 'N° PROY/ COD: AX', font_key: 'label_bold_10' },
    'S12': { texto: 'Fecha Inicio', font_key: 'label_bold_10' },
    'S14': { texto: 'Fecha Término', font_key: 'label_bold_10' },
    'D99': { texto: 'Nombre Supervisor Contratista', font_key: 'label_bold_10' },
    'D102': { texto: 'Nombre Supervisor', font_key: 'label_bold_10' },
    'D104': { texto: 'OBSERVACIONES SUPERVISOR', font_key: 'label_bold_10' },
    'Q104': { texto: 'OBSERVACIONES SUPERVISOR', font_key: 'label_bold_10' },
    'S99': { texto: 'Firma', font_key: 'label_bold_10', alignment_key: 'align_right' },
    'S102': { texto: 'Firma', font_key: 'label_bold_10', alignment_key: 'align_right' },
    'C108': { texto: '* Deben ser completados todos los campos de observacion, siendo responsabilidad del Supervisor Despliegue y Personal de SCM', font_key: 'footnote_10' },
    'C109': { texto: '** Los campos SI (Aceptado), NO (Rechazado) y NA (No aplica) deben ser marcados con una "X".', font_key: 'footnote_10' },
    'C110': { texto: 'Versión 2016-09', font_key: 'footnote_10' }
  },
  celdas_a_combinar_base: [
    'B6:AG6', 'T8:AE8', 'E10:P10', 'T10:Y10', 'Z10:AE10', 'E12:P12', 'T12:Y12', 'Z12:AE12',
    'E14:P14', 'T14:Y14', 'Z14:AE14', 'D16:G16', 'H16:AE16', 'E18:AE18',
    'D99:G99', 'H99:R99', 'T99:AD99', 'E102:R102', 'T102:AD102', 'D104:G104', 'Q104:T104',
    'D105:P105', 'Q105:AE105', 'D106:P106', 'Q106:AE106', 'D107:P107', 'Q107:AE107',
    'C108:AF108', 'C109:AF109', 'C110:AF110'
  ],
  celda_titulo_acta_base: "B6",
  filas_por_bloque_plantilla: 113,
  max_fila_contenido_bloque_base: 110,
  // Dimensiones específicas para fotos en hojas de conformidad, usadas como override en sheet_conformidad.py
  // Estas serán sobrescritas por PHOTO_PIXEL_WIDTH y PHOTO_PIXEL_HEIGHT de PHOTO_CONFIG
  CONFORMIDAD_PHOTO_PIXEL_WIDTH: null, // Será ajustado
  CONFORMIDAD_PHOTO_PIXEL_HEIGHT: null, // Será ajustado
};

// Consolidar todas las configuraciones de estilo y diseño
const styleAndLayoutConfig = {
  RUTA_LOGO_GTD,
  NOMBRE_ARCHIVO_DATOS_PRINCIPAL,
  NOMBRE_ARCHIVO_COMENTARIOS,
  FOLDER_TO_TITLE_MAP,
  MAIN_JSON_CELL_MAP_CONFORMIDAD,
  PHOTO_CONFIG,
  COMMON_FONTS: COMMON_FONTS_EXCELJS, // Renombrado para claridad y evitar confusión con los de Python
  COMMON_ALIGNMENTS: COMMON_ALIGNMENTS_EXCELJS, // Renombrado
  BORDER_STYLES: BORDER_STYLES_EXCELJS, // Renombrado
  CONFORMIDAD_SHEET_CONFIG,
};

// Ajustar CONFORMIDAD_PHOTO_PIXEL_WIDTH y CONFORMIDAD_PHOTO_PIXEL_HEIGHT
// para que apunten a los valores calculados en PHOTO_CONFIG
styleAndLayoutConfig.CONFORMIDAD_SHEET_CONFIG.CONFORMIDAD_PHOTO_PIXEL_WIDTH = styleAndLayoutConfig.PHOTO_CONFIG.PHOTO_PIXEL_WIDTH;
styleAndLayoutConfig.CONFORMIDAD_SHEET_CONFIG.CONFORMIDAD_PHOTO_PIXEL_HEIGHT = styleAndLayoutConfig.PHOTO_CONFIG.PHOTO_PIXEL_HEIGHT;


export default styleAndLayoutConfig;