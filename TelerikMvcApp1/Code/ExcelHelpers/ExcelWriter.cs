using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CB.Excel {
	class ExcelWriter {
		MemoryStream _stream;
		SpreadsheetDocument _document;
		Sheet _sheet;
		SheetData _sheetdata;
		Dictionary<string, int> _strings;
		int _maxstringid;
		Row _row;
		uint _rowIndex = 0;
		int _colIndex;


		public void StartNewExcelDocument() {
			_stream = new MemoryStream();
			_document = SpreadsheetDocument.Create(_stream, SpreadsheetDocumentType.Workbook);
			_document.AddWorkbookPart();
			_document.WorkbookPart.Workbook = new Workbook();
			_document.WorkbookPart.Workbook.Sheets = new Sheets();

			_strings = new Dictionary<string, int>();
			_maxstringid = 0;
		}

		public void CreateNewWorksheet(string name) {
			//add worksheetpart to exsting workbookpart then add new worksheet to worksheetpart
			WorksheetPart worksheetPart = _document.WorkbookPart.AddNewPart<WorksheetPart>();
			_sheetdata = new SheetData();
			worksheetPart.Worksheet = new Worksheet(_sheetdata);

			Sheets sheets = _document.WorkbookPart.Workbook.GetFirstChild<Sheets>();
			uint massheetid = 1;
			if (sheets.Elements<Sheet>().Count() > 0) {
				massheetid = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
			}

			_sheet = new Sheet() {
				Id = _document.WorkbookPart.GetIdOfPart(worksheetPart),
				SheetId = massheetid,
				Name = name
			};

			sheets.Append(_sheet);			
		}

		public byte[] GetCompletedExcelDocument() {
			StoreSharedStrings();
			_document.WorkbookPart.Workbook.Save();
			_document.Close();

			return _stream.ToArray();
		}

		/// <summary>
		/// Converts a column number to column name (i.e. A, B, C..., AA, AB...)
		/// </summary>
		/// <param name="columnIndex">Index of the column</param>
		/// <returns>Column name</returns>
		private string ColumnNameFromIndex(int columnIndex) {
			int remainder;
			string columnName = "";

			columnIndex++;
			while (columnIndex > 0) {
				remainder = (columnIndex - 1) % 26;
				columnName = System.Convert.ToChar(65 + remainder).ToString() + columnName;
				columnIndex = ((columnIndex - remainder) / 26);
			}

			return columnName;
		}


		//private Cell CreateTextCell(string header, UInt32 index,string text) {
		//	var cell = new Cell {
		//		DataType = CellValues.InlineString,
		//		CellReference = header + index
		//	};

		//	var istring = new InlineString();
		//	var t = new Text { Text = text };
		//	istring.AppendChild(t);
		//	cell.AppendChild(istring);
		//	return cell;
		//}

		private Cell CreateSharedTextCell(string header, UInt32 index, string text) {
			int idx = IndexOfSharedString(text);

			var cell = new Cell {
				DataType = CellValues.SharedString,
				CellReference = header + index,
				CellValue = new CellValue(idx.ToString())
			};
			return cell;
		}

		private int IndexOfSharedString(string value) {
			if (_strings.ContainsKey(value)) {
				return _strings[value];
			}
			int id = _maxstringid++;
			_strings.Add(value, id);
			return id;
		}

		private void StoreSharedStrings() {
			var sharedStringTablePart = _document.WorkbookPart.AddNewPart<SharedStringTablePart>();
			var tab = new SharedStringTable();
			sharedStringTablePart.SharedStringTable = tab;

			List<string> orderstringlist = _strings.ToList().OrderBy(x => x.Value).Select(x => x.Key).ToList();

			foreach(string s in orderstringlist){
				tab.AppendChild( new SharedStringItem( new Text(s)));
			}
		}


		public void NewRow() {
			_row = new Row { RowIndex = ++_rowIndex };
			_sheetdata.AppendChild(_row);
			_colIndex = 0;
		}
		
		public void AddCell(int index, string value) {
			if (index < _colIndex) {
				throw new ApplicationException("Cannot reverse column indexing.");
			}
			_colIndex = index;
			_row.AppendChild(CreateSharedTextCell(ColumnNameFromIndex(index), _rowIndex, value));
		}

		public void AddCell(string value) {
			_row.AppendChild(CreateSharedTextCell(ColumnNameFromIndex(_colIndex++), _rowIndex, value));
		}

		public void SkipCell() {
			_colIndex++;
		}
	}
}