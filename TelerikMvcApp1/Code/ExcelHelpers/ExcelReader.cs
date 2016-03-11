using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CB.Excel {
	public class SLExcelStatus {
		public string Message { get; set; }
		public bool Success {
			get { return string.IsNullOrWhiteSpace(Message); }
		}
	}

	public class SLExcelDataRow {
		public int LineNum { get; set; }
		public int Id { get; set; }
		public bool State { get; set; }
		public string Marker { get; set; }

		public List<string> CellData { get; set; }
	}
	
	public class SLExcelData {
		/// <summary>Errors when reading the Excel file</summary>
		public SLExcelStatus Status { get; set; }
		/// <summary>configuration information of the columns in the Excel sheet</summary>
		public Columns ColumnConfigurations { get; set; }
		/// <summary>the data in the first row of the Excel sheet</summary>
		public List<string> Headers { get; set; }
		/// <summary>keeps the data for the rest of the rows of the Excel sheet</summary>
		public List<SLExcelDataRow> DataRows { get; set; }
		/// <summary>keeps the name of the Excel sheet</summary>
		public string SheetName { get; set; }
		public SLExcelData() {
			Status = new SLExcelStatus();
			Headers = new List<string>();
			DataRows = new List<SLExcelDataRow>();
		}
	}
	
	public class SLExcelReader {
		Dictionary<int, string> _sharedstringhash;

		
		private string GetColumnName(string cellReference) {
			var regex = new Regex("[A-Za-z]+");
			var match = regex.Match(cellReference);
	
			return match.Value;
		}

		/// <summary>
		/// Get the index for the given excel column name
		/// </summary>
		/// <param name="columnName"></param>
		/// <returns></returns>
		private int ConvertColumnNameToNumber(string columnName) {
			var alpha = new Regex("^[A-Z]+$");
			if (!alpha.IsMatch(columnName)) throw new ArgumentException();
	
			char[] colLetters = columnName.ToCharArray();
			Array.Reverse(colLetters);
	
			var convertedValue = 0;
			for (int i = 0; i < colLetters.Length; i++) {
				char letter = colLetters[i];
				// ASCII 'A' = 65
				int current = i == 0 ? letter - 65 : letter - 64;
				convertedValue += current * (int)Math.Pow(26, i);
			}
	
			return convertedValue;
		}

		private IEnumerator<Cell> GetExcelCellEnumerator(Row row) {
			int currentCount = 0;
			foreach (Cell cell in row.Descendants<Cell>()) {
				string columnName = GetColumnName(cell.CellReference);
	
				int currentColumnIndex = ConvertColumnNameToNumber(columnName);
	
				for (; currentCount < currentColumnIndex; currentCount++) {
					var emptycell = new Cell() {
						DataType = null, CellValue = new CellValue(string.Empty)
					};
					yield return emptycell;
				}
	
				yield return cell;
				currentCount++;
			}
		}
	
		private string ReadExcelCell(Cell cell) {
			var cellValue = cell.CellValue;
			var text = (cellValue == null) ? cell.InnerText : cellValue.Text;
			if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString)) {
				text = _sharedstringhash[Convert.ToInt32(cell.CellValue.Text)];
			}
	
			return (text ?? string.Empty).Trim();
		}
	


		public SLExcelData ReadExcel(string filename) {
			var data = new SLExcelData();
	
			//// Check if the file is excel
			//if (file.ContentLength <= 0) {
			//	data.Status.Message = "You uploaded an empty file";
			//	return data;
			//}
	
			//if (file.ContentType 
			//	!= "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
			//	data.Status.Message
			//	= "Please upload a valid excel file of version 2007 and above";
			//	return data;
			//}
	
			// Open the excel document
			WorkbookPart workbookPart;
			SpreadsheetDocument document;
			List<Row> rows;
			try {
				document = SpreadsheetDocument.Open(filename, false);
				workbookPart = document.WorkbookPart;
	
				var sheets = workbookPart.Workbook.Descendants<Sheet>();
				var sheet = sheets.First();
				data.SheetName = sheet.Name;
	
				var workSheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id)).Worksheet;
				var columns = workSheet.Descendants<Columns>().FirstOrDefault();
				data.ColumnConfigurations = columns;
	
				var sheetData = workSheet.Elements<SheetData>().First();
				rows = sheetData.Elements<Row>().ToList();
			}catch (Exception e) {
				data.Status.Message = "Unable to open the file. "+e.Message;
				return data;
			}

			//as a way of compression any strings that are the same are referenced rather than repeated. Build a quick way of getting this shared strings
			_sharedstringhash = new Dictionary<int, string>();
			if (workbookPart.SharedStringTablePart != null) {
				var sharedlist = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>();
				int sli = 0;
				foreach (SharedStringItem ssi in sharedlist) {
					_sharedstringhash.Add(sli, ssi.InnerText);
					sli++;
				}
			}
			

	
			// Read the header
			if (rows.Count > 0) {
				var row = rows[0];
				var cellEnumerator = GetExcelCellEnumerator(row);
				while (cellEnumerator.MoveNext()) {
					var cell = cellEnumerator.Current;
					var text = ReadExcelCell(cell).Trim();
					data.Headers.Add(text);
				}
			}
	
			// Read the sheet data
			if (rows.Count > 1) {
				for (var i = 1; i < rows.Count; i++) {
					var dataRow = new SLExcelDataRow() { LineNum = i, Id=0, State = false, Marker = string.Empty, CellData = new List<string>() };
					data.DataRows.Add(dataRow);
					var row = rows[i];
					var cellEnumerator = GetExcelCellEnumerator(row);
					while (cellEnumerator.MoveNext()) {
						var cell = cellEnumerator.Current;
						var text = ReadExcelCell(cell).Trim();
						dataRow.CellData.Add(text);
					}
				}
			}

			document.Close();

			return data;
		}
	}
}

