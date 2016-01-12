using System;
using System.Collections.Generic;
using System.IO;
using Excel;
using System.Data;
using Aspose.Cells;
using System.Reflection;
using System.Linq;
using System.Collections;

namespace MergeExcel
{
	class MainClass
	{
		private const int MaxColNumber = 27;

		private static Dictionary<String, List<List<String>>> _problemas;
		private static Dictionary<String, List<List<String>>> _riesgos;


		public static void Main (string[] args)
		{

			try {
				//args = new String[] {"/Users/PabloZaiden/Desktop/test/"};

				if (args.Length != 1) {
					Console.Error.WriteLine ("Usage: MergeExcel <DIR>");
					return;
				}

				_riesgos = new Dictionary<string, List<List<string>>>();
				_problemas = new Dictionary<string, List<List<string>>>();

				var dir = new DirectoryInfo (args [0]);

				ReadExcelFilesRecursive (dir);

				var outputFileNameProblemas = "Problemas_" + DateTime.Today.ToString("yy-MM-dd") + ".xlsx";
				var outputFileNameRiesgos = "Riesgos_" + DateTime.Today.ToString("yy-MM-dd") + ".xlsx";

				WriteExcelFile(_problemas, outputFileNameProblemas);
				WriteExcelFile(_riesgos, outputFileNameRiesgos);

			} catch (Exception e) {
				Console.Error.WriteLine (e);
			}

			Console.WriteLine ("Proceso Finalizado. Presione cualquier tecla para continuar...");
			Console.ReadKey ();
		}

		static void WriteExcelFile (Dictionary<string, List<List<string>>> data, string fileName)
		{
			Console.WriteLine ("Escribiendo: " + fileName);
			using (Workbook book = new Aspose.Cells.Workbook ()) {
				foreach (var key in data.Keys) {
					var sheet = book.Worksheets.Add (key);
		
					var rows = data [key];
					for (int i = 0; i < rows.Count; i++) {
					
						for (int j = 0; j < rows [i].Count; j++) {
							sheet.Cells.Rows [i] [j].Value = rows [i] [j];
							if (i == 0) {
								var style = sheet.Cells.Rows [i] [j].GetStyle ();
								style.Font.IsBold = true;
								sheet.Cells.Rows [i] [j].SetStyle (style);
							}
						}
					}
				}

				book.Save (fileName);
			}
		}

		static void ReadExcelFilesRecursive (DirectoryInfo dir)
		{
			ReadExcelFiles2 (dir);
			foreach (var subdir in dir.GetDirectories()) {
				ReadExcelFilesRecursive (subdir);
			}
		}

		/*
		static void ReadExcelFiles (DirectoryInfo dir)
		{
			foreach (var f in dir.GetFiles ("*.xlsx")) {
				var nombre = f.Name.ToLowerInvariant ();
				if ((nombre.Contains("riesgo") || nombre.Contains("problema")) && !nombre.Contains ("$") ) {
					Console.WriteLine ("Leyendo: " + f.FullName);
					using (var stream = f.OpenRead ()) {

						var reader = ExcelReaderFactory.CreateOpenXmlReader (stream);
						Dictionary<String, List<List<String>>> dic;
						if (f.Name.ToLowerInvariant ().Contains ("riesgo")) {
							dic = _riesgos;
						} else {
							dic = _problemas;
						}
						ProcessTable (f.FullName, reader, dic);
					}
				}
			}
		}
		*/

		static string GetResultTableName (IExcelDataReader reader, int i)
		{
			
			var fields = reader.GetType ().GetRuntimeFields ();
			var workbookField = fields.FirstOrDefault(f => f.Name.ToLowerInvariant().Contains("workbook")); 

			var workBook = workbookField.GetValue (reader);

			var workbookType = workBook.GetType ();
			var sheets = (IList)workbookType.GetProperty ("Sheets").GetValue (workBook);

			var sheet = sheets [i];
			string name = (String)sheet.GetType ().GetProperty ("Name").GetValue (sheet);
			//return "";
			return name;
		}

		/*
		static void ProcessTable (String name, IExcelDataReader reader, Dictionary<String, List<List<string>>> dic)
		{
			for (int i = 0; i < reader.ResultsCount; i++) {
				
				String tableName = GetResultTableName (reader, i);

				reader.Read ();
				int rowNum = 0;

				if (Char.IsNumber (tableName[0])) {
					List<List<String>> lista;
					if (dic.ContainsKey (tableName)) {
						lista = dic [tableName];
					} else {
						lista = new List<List<string>> ();
						dic.Add (tableName, lista);
					}

					if (lista.Count == 0) {
						//leer cabecera
						//while (rowNum < 1) {
						//	reader.Read ();
						//	rowNum++;
						//
					
						List<String> header = new List<string> ();
						header.Add ("Archivo");
						for (int j = 0; j < Math.Min(MaxColNumber, reader.FieldCount); j++) {
							
							var value = reader.GetValue (j);
							header.Add (IsEmpty(value) ? "" : value.ToString());
						}
						lista.Add (header);
					}

					//while (rowNum < 6) {
					//	reader.Read ();
					//	rowNum++;
					//}

					while (!IsEmpty (reader.GetValue(1))) {
						List<String> elem = new List<string> ();
						elem.Add (name);
						for (int j = 0; j < Math.Min(MaxColNumber, reader.FieldCount); j++) {
							var value = reader.GetValue (j);
							elem.Add (IsEmpty(value) ? "" : value.ToString());
						}
						lista.Add (elem);
						reader.Read ();
						rowNum++;
					}

				}
				reader.NextResult ();
			}
		}
		*/

		static void ReadExcelFiles2 (DirectoryInfo dir)
		{
			foreach (var f in dir.GetFiles ("*.xlsx")) {
				var nombre = f.Name.ToLowerInvariant ();

				if ((nombre.Contains("riesgo") || nombre.Contains("problema")) && !nombre.Contains ("$") ) {
						
					Console.WriteLine ("Leyendo: " + f.FullName);
					using (var stream = f.OpenRead ()) {
					

						var dataset = ExcelReaderFactory.CreateOpenXmlReader (stream).AsDataSet ();
						Dictionary<String, List<List<String>>> dic;
						if (f.Name.ToLowerInvariant ().Contains ("riesgo")) {
							dic = _riesgos;
						} else {
							dic = _problemas;
						}
						ProcessTable2 (f.FullName, dataset, dic);
					}
				}
			}
		}

		static void ProcessTable2 (String name, DataSet dataset, Dictionary<String, List<List<string>>> dic)
		{
			foreach (DataTable table in dataset.Tables) {
				if (Char.IsNumber (table.TableName [0])) {

					List<List<String>> lista;
					if (dic.ContainsKey (table.TableName)) {
						lista = dic [table.TableName];
					} else {
						lista = new List<List<string>> ();
						dic.Add (table.TableName, lista);
					}

					if (lista.Count == 0) {
						//leer cabecera
						var row = table.Rows[5];
						List<String> header = new List<string> ();
						header.Add ("Archivo");
						for (int i = 0; i < MaxColNumber; i++) {
							header.Add (row [i].ToString ());
						}
						lista.Add (header);
					}

					int rowNum = 6;
					while (!IsEmpty (table.Rows [rowNum] [1])) {
						var row = table.Rows [rowNum];
						List<String> elem = new List<string> ();
						elem.Add (name);
						for (int i = 0; i < MaxColNumber; i++) {
							elem.Add (row [i].ToString ());
						}
						lista.Add (elem);
						rowNum++;
					}

				}
			}
		}

		static bool IsEmpty (object obj)
		{
			if (obj is DBNull) {
				return true;
			} else  if (obj is String) {
				return String.IsNullOrWhiteSpace ((String)obj);
			} else {
				return obj == null;
			}
		}
	}
}
