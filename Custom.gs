//get all files xls from certain folder in drive and convert to spreadsheet with same name and after delete xls
function convert(folderId) {

	var folderIncoming = DriveApp.getFolderById("1Bjr9xk5C11liHJj7_bR3IRQQK5lt3b9-");
	var files = folderIncoming.getFilesByType(MimeType.MICROSOFT_EXCEL_LEGACY);

	while (files.hasNext()) {
		var source = files.next();
		var sourceId = source.getId();
		var fileName = source.getName().replace('.xls', '');

		var file = {
			title: fileName,
		};
 
		/* delete files .xls with name Уценка, уценка and Бланк */
		var regExp = /.*(Уценка|уценка|Бланк).*/g
		if (fileName.match(regExp)) {
			Drive.Files.trash(sourceId)
		}
		/* */

		file = Drive.Files.copy(file, sourceId, {
			convert: true
		});
      
        replaceSpecialSymbols();

		var SpreadsheetID = file.getId();
		var SheetName = "TDSheet";
		var ss = SpreadsheetApp.openById(SpreadsheetID)
		var sheet = ss.getSheetByName(SheetName);
		var data = sheet.getRange("A:Z");
		var values = data.getValues();
		var numRows = values[0].length;
		var numCols = values[0].length;

		for (var col = numCols - 1; col > 0; col--) { // count down over columns   
			for (var row = 0; row < numRows; row++) { // count up over rows
				switch (values[row][col]) { // examine cell contents
					case "Сезон":
					case "Ростовка":
					case "Номенклатура":
					case "Ростовка":
					case "Цвет":
					case "Размер колес":
					case "Материал рамы":
					case "Количество скоростей":
					case "Итоговая скидка (%)":
					case "Итоговая цена (RUB)":
					case "Заказ Кол-во (шт.)":
					case "Сумма заказа (руб.)":
					case "Штрихкод":
					case "Вес брутто (кг)":
					case "Объем (м3)":
					case "Итого вес брутто (кг)":
					case "Итого объем (м3)":
					case "Примечания":
					case "Скидка по условиям продаж":
					case "Кратность заказа (шт.)":
					case "Скидка по условиям продаж (%)":
					case "Скидка за заказ кратно упаковкам (%)":
					case "Номенклатура":
					case "МРРЦ (RUB)":
					case "Цена (RUB)": // in these cases...
						sheet.deleteColumn(col + 1); // delete column in sheet (1-based)     
						continue; // continue with next column
						break; // can't get here, but good practice

				}
			}

		}

		sheet.deleteColumn(1); //удаление первого столбца (Сезон)
		sheet.deleteRows(1, 10); //удаление с 1 по 10 строки


		deleteSheets();
		mergeSheets();

		file = Drive.Files.trash(sourceId) //удаление исходных файлов .xsl


	}
	finalClear();
}

function deleteSheets() {
	var ss = SpreadsheetApp.openById("1RUbVaM209dcDe4B2R35jwQ9iy6yNevU_H7QpBGR41Hg");
	var sheets = ss.getSheets();
	for (i = 0; i < sheets.length; i++) {
		switch (sheets[i].getSheetName()) {
			case "1":
			case "2":
			case "3":
			case "4":
				break;
			default:
				ss.deleteSheet(sheets[i]);
		}
	}
}

function mergeSheets() {

	/* Retrieve the desired folder */
	var myFolder = DriveApp.getFolderById("1Bjr9xk5C11liHJj7_bR3IRQQK5lt3b9-");

	/* Get all spreadsheets that resided on that folder */
	var spreadSheets = myFolder.getFilesByType("application/vnd.google-apps.spreadsheet");

	/* Create the new spreadsheet that you store other sheets */
	var newSpreadSheet = SpreadsheetApp.openById("1RUbVaM209dcDe4B2R35jwQ9iy6yNevU_H7QpBGR41Hg");

	/* Iterate over the spreadsheets over the folder */
	while (spreadSheets.hasNext()) {

		var sheet = spreadSheets.next();

		/* Open the spreadsheet */
		var spreadSheet = SpreadsheetApp.openById(sheet.getId());

		/* Get all its sheets */
		for (var y in spreadSheet.getSheets()) {

			/* Copy the sheet to the new merged Spread Sheet */
			spreadSheet.getSheets()[y].copyTo(newSpreadSheet);
			SpreadsheetApp.flush(); //update all tabs


		}

	}

}

function finalClear() {

	/* Retrieve the desired folder */
	var myFolder = DriveApp.getFolderById("1Bjr9xk5C11liHJj7_bR3IRQQK5lt3b9-");

	/* Get all spreadsheets that resided on that folder */
	var spreadSheets = myFolder.getFilesByType("application/vnd.google-apps.spreadsheet");

	/* Iterate over the spreadsheets over the folder */
	while (spreadSheets.hasNext()) {

		var source = spreadSheets.next();
		var sourceId = source.getId();
		var fileName = source.getName();

		/* delete files with name Подольск */
		var regExp = /.*(Подольск).*/g
		if (fileName.match(regExp)) {
			Drive.Files.trash(sourceId)
		}
		/* */

	}

}

function replaceSpecialSymbols() {
  
  var myFolder = DriveApp.getFolderById("1Bjr9xk5C11liHJj7_bR3IRQQK5lt3b9-");
  var spreadSheets = myFolder.getFilesByType("application/vnd.google-apps.spreadsheet");
  
  while (spreadSheets.hasNext()) {

		var source = spreadSheets.next();
		var sourceId = source.getId();
		var fileName = source.getName();

		/* delete files with name Подольск */
		var regExp = /.*(Подольск).*/g
		if (fileName.match(regExp)) {
			
          var sheet = SpreadsheetApp.openById(sourceId);
          var range = sheet.getRange("A:Z"); /* range you want to modify */
          var data  = range.getValues();
          
          for (var row = 0; row < data.length; row++) {
            for (var col = 0; col < data[row].length; col++) {
              data[row][col] = (data[row][col]).toString().replace(/>/g, ' ');
            }
          }
          range.setValues(data);
          
          
		}

	}
 
}
