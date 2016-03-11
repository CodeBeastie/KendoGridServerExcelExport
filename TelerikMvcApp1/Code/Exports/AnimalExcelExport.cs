using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using TelerikMvcApp1.Data;
using TelerikMvcApp1.ViewModels;

namespace CB.Excel {
	public class AnimalExcelExport {

		static public byte[] GetExcelFile(  IEnumerable<AnimalViewModel> dataToExport) {



			ExcelWriter writer = new ExcelWriter();
			writer.StartNewExcelDocument();
			writer.CreateNewWorksheet("TagAttributes");
			writer.NewRow();
			writer.AddCell("Id");
			writer.AddCell("Name");
			writer.AddCell("AnimalType");
			writer.AddCell("InZoo");
			writer.AddCell("Age");


			var en = dataToExport.GetEnumerator();
			AnimalViewModel rowdata;
			while (en.MoveNext()) {
				rowdata = en.Current;
				if (rowdata != null) {
					writer.NewRow();
					writer.AddCell(rowdata.Id.ToString());
					writer.AddCell(rowdata.Name);
					writer.AddCell(rowdata.AnimalType);
					writer.AddCell(rowdata.InZoo.ToString());
					writer.AddCell(rowdata.Age.ToString());
				}
			}


			return writer.GetCompletedExcelDocument();
		}
		
	}
}
