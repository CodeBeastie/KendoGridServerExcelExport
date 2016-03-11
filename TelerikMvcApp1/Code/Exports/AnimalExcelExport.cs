using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using TelerikMvcApp1.Data;
using TelerikMvcApp1.ViewModels;

namespace CB.Excel {
	public class AnimalExcelExport {

		static public byte[] GetExcelFile(  IEnumerable<AnimalViewModel> dataToExport) {


			//create excel using the same fields shown on the grid
			ExcelWriter writer = new ExcelWriter();
			writer.StartNewExcelDocument();
			writer.CreateNewWorksheet("TagAttributes");
			writer.NewRow();
			writer.AddCell("Id");
			writer.AddCell("Name");
			writer.AddCell("AnimalType");
			writer.AddCell("InZoo");
			writer.AddCell("Age");
			//TagAttributeEditorViewModel vm = service.GetTagAttributeEditorViewmodel(projectId, userId);
			//foreach (TagAttributeColumnViewModel col in vm.AttrColumns) {
			//	writer.AddCell(col.Title);
			//	if (col.PropertyName.StartsWith("L")) {
			//		col.PropertyName = "S" + col.PropertyName.Substring(1);		//want the string not the list id
			//	}
			//}


			//Type t = typeof(AnimalViewModel);
			//var properties = t.GetProperties().ToDictionary(x => x.Name, x => x);

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
					//foreach (TagAttributeColumnViewModel col in vm.AttrColumns) {
					//	PropertyInfo prop = properties[col.PropertyName];
					//	if (prop.GetValue(rowdata) == null) {
					//		writer.SkipCell();
					//	} else {
					//		writer.AddCell(prop.GetValue(rowdata).ToString());
					//	}
					//}
				}
			}


			return writer.GetCompletedExcelDocument();
		}
		
	}
}
