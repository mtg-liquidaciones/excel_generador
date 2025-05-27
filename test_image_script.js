// test_image_script.js
import ExcelJS from 'exceljs';
import fs from 'fs/promises';
import path from 'path';

async function runImageTest() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Test de Imagen');
  // CAMBIA ESTA RUTA A TU IMAGEN PNG DE PRUEBA DE 50x50
  const imagePath = "C:/Users/MAOC/Desktop/logogtd.png"; // O la ruta a tu logogtd.png

  console.log(`--- Test Script Iniciado ---`);
  console.log(`Intentando añadir imagen desde: ${imagePath}`);

  try {
    await fs.access(imagePath);
    console.log(`Acceso al archivo confirmado: ${imagePath}`);

    const imageBuffer = await fs.readFile(imagePath);
    console.log(`Archivo leído en buffer, tamaño: ${imageBuffer.length} bytes`);

    const extension = path.extname(imagePath).substring(1).toLowerCase() || 'png';
    console.log(`Extensión deducida/usada: ${extension}`);

    const imageId = workbook.addImage({
      buffer: imageBuffer,
      extension: extension, // 'png' o 'jpeg'
    });
    console.log(`Imagen añadida al libro (workbook). ImageId: ${JSON.stringify(imageId)}`);
    console.log(`Detalles del ImageId: ID=<span class="math-inline">\{imageId\.id\}, Type\=</span>{imageId.type}`);


    // Colocar en celda B2 con dimensiones explícitas
    sheet.addImage(imageId, {
      tl: { col: 1, row: 1 }, // Corresponde a la celda B2 (0-indexed)
      ext: { width: 50, height: 50 } // Usa las dimensiones de tu imagen de prueba
    });
    console.log('Imagen colocada en la hoja en B2 con dimensiones 50x50.');

    // También intenta colocarla con un rango, que usa las dimensiones originales de la imagen
    // const imageId2 = workbook.addImage({ buffer: imageBuffer, extension: extension });
    // sheet.addImage(imageId2, 'D2:F8');
    // console.log('Imagen colocada en la hoja en D2:F8 (tamaño original).');


    await workbook.xlsx.writeFile('output_test_imagen.xlsx');
    console.log('Archivo Excel de prueba "output_test_imagen.xlsx" guardado en la raíz del proyecto.');
    console.log('--- Test Script Finalizado ---');

  } catch (error) {
    console.error('ERROR DURANTE EL TEST DE IMAGEN:', error);
  }
}

runImageTest();