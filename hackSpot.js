const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

async function clear() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('../hack_excel/0.xlsx');
    const worksheet = workbook.getWorksheet('Spot');
    
    // Удаление первых пяти строк
    worksheet.spliceRows(1, 5);

    // Добавление 4x новых строк в начало таблицы
    worksheet.insertRow(1, []);
    worksheet.insertRow(1, []);
    worksheet.insertRow(1, []);
    worksheet.insertRow(1, []);

    // Удаление колонок учитывая текущие изменения
    worksheet.spliceColumns(1, 1); // Удаляем колонку A
    worksheet.spliceColumns(4, 1); // Удаляем колонку D, которая теперь представляет собой бывшую E
    worksheet.spliceColumns(16, 1); // Удаляем колонку R

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

    await workbook.xlsx.writeFile('../hack_excel/1.xlsx');
    return { workbook, worksheet }; // Возвращаем workbook и worksheet для дальнейшего использования
}

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

// Проверяем совпадение названий листов и копируем данные, если совпадают
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

	// Устанавливаем фиксированную ширину для всех столбцов
	const fixedWidth = 15; // Здесь задаем фиксированную ширину в символах

	targetSheet.columns.forEach((column, colNumber) => {
		 column.width = fixedWidth;
	});

	// Сохраняем изменения в целевом файле
	await targetWorkbook.xlsx.writeFile('../hack_excel/result/template.xlsx');
}



async function hackSpot() {
    const { workbook, worksheet } = await clear();
    await draw(worksheet); // Здесь мы используем уже модифицированный worksheet из `clear`
    await workbook.xlsx.writeFile('../hack_excel/result/spot.xlsx'); // Сохраняем окончательные изменения

    // Проверяем совпадение названий листов и копируем данные, если совпадают
    const sourceWorkbook = new ExcelJS.Workbook();
    await sourceWorkbook.xlsx.readFile('../hack_excel/result/spot.xlsx');
    const targetWorkbook = new ExcelJS.Workbook();
    await targetWorkbook.xlsx.readFile('../hack_excel/result/template.xlsx');

    const sourceSheetName = 'Spot';
    const targetSheetName = 'Spot';

    if (sourceWorkbook.getWorksheet(sourceSheetName) && targetWorkbook.getWorksheet(targetSheetName)) {
        await copyData(sourceWorkbook, sourceSheetName, targetWorkbook, targetSheetName);
        console.log('Копирование завершено успешно');
    } else {
        console.log('Листы с заданными именами не найдены');
    }
}

hackSpot().then(() => console.log('Ты просто Джедай, братан 😎'));
