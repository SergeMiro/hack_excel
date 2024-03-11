
const ExcelJS = require('exceljs');

async function removeColumns() {
	const workbook = new ExcelJS.Workbook();
	await workbook.xlsx.readFile('../hack_excel/CIU Table.xlsx');
	const worksheet = workbook.getWorksheet('Spot');
	
	// Удаление первых пяти строк
	worksheet.spliceRows(1, 5);

	// Добавление двух новых строк в начало таблицы
	worksheet.insertRow(1, []);
	worksheet.insertRow(1, []);

	// Удаление колонок учитывая текущие изменения
	worksheet.spliceColumns(1, 1); // Удаляем колонку A
	worksheet.spliceColumns(4, 1); // Удаляем колонку D, которая теперь представляет собой бывшую E

	// Очистка содержимого в первых двух новых строках (если требуется)
	for (let col = 1; col <= worksheet.columnCount; col++) {
	  worksheet.getRow(1).getCell(col).value = null;
	  worksheet.getRow(2).getCell(col).value = null;
	}

	await workbook.xlsx.writeFile('../hack_excel/CIU_Table_modified.xlsx');
	return { workbook, worksheet }; // Возвращаем workbook и worksheet для дальнейшего использования
}















//////////////////////////////////////////////////////////////////////////////////
async function mergeCells(worksheet) {
  const mergeRanges = [
    'A1:A2', 'B1:B2', 'C1:C2', 'D1:D2', 'E1:F1', 'F1:G1', 'H1:K1', 'L1:L2', 'M1:M2', 'N1:N2', 'O1:O2'
  ];

  mergeRanges.forEach(range => {
    try {
      worksheet.mergeCells(range);
    } catch (error) {
      console.log(`Ошибка при объединении диапазона ${range}: ${error}`);
    }
  });

  mergeRanges.forEach(range => {
    worksheet.getCell(range.split(':')[0]).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
  });
}  

//////////////////////////////////////////////////////////////////////////////
async function hackSpot() {
  const { workbook, worksheet } = await removeColumns();
  await mergeCells(worksheet);
  await workbook.xlsx.writeFile('../hack_excel/CIU_Table_modified_final.xlsx'); // Сохраняем окончательные изменения
}
hackSpot().then(() => console.log('Лишние колонки и строки удалены, ячейки объединены 😎'));
