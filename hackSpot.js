
const ExcelJS = require('exceljs');


async function clear() {
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

	// Очистка содержимого в первых двух новых строках
	for (let col = 1; col <= worksheet.columnCount; col++) {
		 worksheet.getRow(1).getCell(col).value = null;
		 worksheet.getRow(2).getCell(col).value = null;
	}

	// Очистка границ всех ячеек на листе
	worksheet.eachRow({ includeEmpty: true }, function(row) {
		row.eachCell({ includeEmpty: true }, function(cell) {
			 cell.border = {
				  top: { style: 'thin', color: { argb: 'FFD3D3D3' } },
				  left: { style: 'thin', color: { argb: 'FFD3D3D3' } },
				  bottom: { style: 'thin', color: { argb: 'FFD3D3D3' } },
				  right: { style: 'thin', color: { argb: 'FFD3D3D3' } }
			 };
		});
  });
  
	await workbook.xlsx.writeFile('../hack_excel/CIU_Table_modified.xlsx');
	return { workbook, worksheet }; // Возвращаем workbook и worksheet для дальнейшего использования
}


//////////////////////////////////////////////////////////////////////////////////
function setBordersForNonEmptyCells(worksheet) {
	worksheet.eachRow({ includeEmpty: true }, function(row) {
		 row.eachCell({ includeEmpty: true }, function(cell) {
			  if (cell.value !== null && cell.value !== '') {
					cell.border = {
						 top: { style: 'thin', color: { argb: 'FF000000' } },
						 left: { style: 'thin', color: { argb: 'FF000000' } },
						 bottom: { style: 'thin', color: { argb: 'FF000000' } },
						 right: { style: 'thin', color: { argb: 'FF000000' } }
					};
			  }
		 });
	});
}

async function draw(worksheet) {
	// Так как `setBordersForNonEmptyCells` уже определена, просто вызываем её с текущим листом
	setBordersForNonEmptyCells(worksheet);
}
 //////////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////////
async function hackSpot() {
	const { workbook, worksheet } = await clear();
	await draw(worksheet); // Здесь мы используем уже модифицированный worksheet из `clear`
	await workbook.xlsx.writeFile('../hack_excel/CIU_Table_modified_final.xlsx'); // Сохраняем окончательные изменения
}
hackSpot().then(() => console.log('Лишние колонки и строки удалены 😎'));

