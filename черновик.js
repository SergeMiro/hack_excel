
async function copyData(sourceWorkbook, sourceSheetName, targetWorkbook, targetSheetName) {
	const sourceSheet = sourceWorkbook.getWorksheet(sourceSheetName);
	const targetSheet = targetWorkbook.getWorksheet(targetSheetName);

	sourceSheet.eachRow(function(sourceRow, rowNumber) {
		 sourceRow.eachCell({ includeEmpty: true }, function(cell, colNumber) {
			  const targetCell = targetSheet.getCell(rowNumber, colNumber);
			  targetCell.value = cell.value;

			  // Копирование стилей ячеек
			  if (cell.style) {
					targetCell.style = Object.assign({}, cell.style);
			  }
		 });
	});

	// Выставляем ширину столбцов
	targetSheet.columns.forEach((column, colNumber) => {
		 let maxWidth = 0;
		 column.eachCell({ includeEmpty: true }, (cell) => {
			  const width = cell.value ? cell.value.toString().length : 0;
			  if (width > maxWidth) {
					maxWidth = width;
			  }
		 });
		 targetSheet.getColumn(colNumber + 1).width = maxWidth + 2; // Добавляем немного дополнительного пространства
	});

	// Сохранение изменений в целевом файле
	await targetWorkbook.xlsx.writeFile('../hack_excel/result/template.xlsx');
}
