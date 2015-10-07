'use strict'

var XLSX = require('xlsx'),
    _ = require('underscore');

module.exports = function parseXlsx(xlsxFile) {

	var workbook = XLSX.readFile(xlsxFile);
	var sheet_name_list = workbook.SheetNames;
	var data = [];

	var colToInt = function(col) {
		var letters = ["", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
		col = col.trim().split('');

		var n = 0;
		for (var i = 0; i < col.length; i++) {
			n *= 26;
			n += letters.indexOf(col[i]); 
		}
		return n;
	};

	var CellCoords = function(cell) {
		cell = cell.split(/([0-9]+)/);
		this.row = parseInt(cell[1]);
		this.column = colToInt(cell[0]);
	};

	var SheetMeta = function(refCell) {
		var dims = refCell.split(':');
		this.firstCol = new CellCoords(dims[0]);
		this.lastCol = new CellCoords(dims[1]);
		this.width =  (this.lastCol.column - this.firstCol.column)+1;
	};

	sheet_name_list.forEach(function(y) { 
		var worksheet = workbook.Sheets[y];
		var sheetData = [];
		data.push(sheetData);
		var sheetMeta = new SheetMeta(worksheet['!ref']);

		var cellCount = 0;
		for (var z in worksheet) {
			if(z[0] === '!' ) continue;
			if(cellCount%sheetMeta.width == 0 ) {
				var row = _.range(6).map(function () { return ' ' })
				sheetData.push(row);
			}
			row[cellCount%sheetMeta.width] = worksheet[z].v;
			cellCount++;
		}
	});

	return data;
};
