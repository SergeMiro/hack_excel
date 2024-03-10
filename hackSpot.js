// const ExcelJS = require('exceljs');

// async function removeColumns() {
//   const workbook = new ExcelJS.Workbook();
//   await workbook.xlsx.readFile('../Excel project/CIU Table.xlsx'); // Укажите правильный путь к файлу
//   const worksheet = workbook.getWorksheet('Spot'); // используйте имя листа
//   // Удаление колонки A (индекс 1)
//   worksheet.spliceColumns(1, 1);
//   // После удаления колонки A, колонка E становится колонкой D (индекс 4)
//   worksheet.spliceColumns(4, 1); // Индекс 4, т.к. колонка E сдвинулась влево после удаления колонки A

//   // Удаляем строки в обратном порядке, чтобы избежать проблемы изменения индексов
//   worksheet.spliceRows(5, 1); // Удаление 5-й строки
//   worksheet.spliceRows(4, 1); // Удаление 4-й строки
//   worksheet.spliceRows(1, 1); // Удаление 1-й строки

// // Перед объединением ячеек "разъединяем" ячейки в первых двух строках
// const totalColumns = worksheet.columnCount; // Предполагаем, что columnCount актуален и отражает общее количество колонок
// for (let row = 1; row <= 2; row++) {
//   for (let col = 1; col <= totalColumns; col++) {
//     // Здесь намеренно пропущен шаг восстановления значения, так как цель - очистить ячейки от текста
//     worksheet.getRow(row).getCell(col).value = null; // Прямо устанавливаем значение каждой ячейки в 'null', очищая её
//   }
//  }
//  await workbook.xlsx.writeFile('../Excel project/CIU_Table_modified.xlsx'); // путь сохранения
// }
// removeColumns().then(() => console.log('Лишние колонки и строки удаленны 😎'));




// // объединение ячеек
// function mergeCells() {
// 	const mergeRanges = [
// 		'A1:A2', // Бывшая B
// 		'B1:B2', // Бывшая C
// 		'C1:C2', // Бывшая D
// 		'D1:E1', // Бывшая E:F, E удалена, F становится E
// 		'F1:G1', // Бывшая G:H
// 		'H1:K1', // Бывшая I:L после сдвига влево
// 		'L1:L2', // Бывшая M
// 		'M1:M2', // Бывшая N
// 		'N1:N2', // Бывшая O
// 		'O1:O2'  // Бывшая P
// 	 ];
// 	mergeRanges.forEach(range => {
// 	  try {
// 		 worksheet.mergeCells(range);
// 	  } catch (error) {
// 		 if (error.message === 'Cannot merge already merged cells') {
// 			console.log(`Ячейки ${range} уже были объединены раньше, но нам море по колена 💦`);
// 		 } else {
// 			throw error; // Если ошибка другого типа, пробрасываем её дальше
// 		 }
// 	  }
// 	});
// 	// Добавление границ ко всем объединенным ячейкам
// 	mergeRanges.forEach(range => {
// 		worksheet.getCell(range.split(':')[0]).border = {
// 		top: { style: 'thin' },
// 		left: { style: 'thin' },
// 		bottom: { style: 'thin' },
// 		right: { style: 'thin' }
// 		};
// 	});
// 	// перезаписываем созраненный файл
// }


// function hackSpot() {
// 	removeColumns()
// 	mergeCells()

// }








/////////////////////////////////////////////




const ExcelJS = require('exceljs');

async function removeColumns() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('../Excel project/CIU Table.xlsx');
  const worksheet = workbook.getWorksheet('Spot');
  
  worksheet.spliceColumns(1, 1);
  worksheet.spliceColumns(4, 1); // Учитывая, что индексы сдвигаются после каждого удаления

  worksheet.spliceRows(5, 1);
  worksheet.spliceRows(4, 1);
  worksheet.spliceRows(1, 1);

  for (let row = 1; row <= 2; row++) {
    for (let col = 1; col <= worksheet.columnCount; col++) {
      worksheet.getRow(row).getCell(col).value = null;
    }
  }

  await workbook.xlsx.writeFile('../Excel project/CIU_Table_modified.xlsx');
  return { workbook, worksheet }; // Возвращаем workbook и worksheet для дальнейшего использования
}

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

async function hackSpot() {
  const { workbook, worksheet } = await removeColumns();
  await mergeCells(worksheet);
  await workbook.xlsx.writeFile('../Excel project/CIU_Table_modified_final.xlsx'); // Сохраняем окончательные изменения
}

hackSpot().then(() => console.log('Лишние колонки и строки удалены, ячейки объединены 😎'));
