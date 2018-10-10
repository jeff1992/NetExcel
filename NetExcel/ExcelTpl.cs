using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetExcel
{
	public class ExcelTpl
    {
		/// <summary>
		/// 
		/// </summary>
		/// <param name="templateFile">template file</param>
		public ExcelTpl(string templateFile)
		{
			this.templateFile = templateFile;
			CheckFile();
		}


		ExcelPackage excelPacket;
		string templateFile;

		/// <summary>
		/// Get the key value dictionary for render work
		/// </summary>
		public Dictionary<string, object> KeyValues { get; } = new Dictionary<string, object>();

		void CheckFile()
		{
			if (!System.IO.File.Exists(templateFile))
				throw new Exception($"File \"{templateFile}\" does not exist");
		}

		/// <summary>
		/// Render template file and save
		/// </summary>
		/// <param name="outputFile">full file name to output</param>
		public void RenderAndSave(string outputFile)
		{
			CheckFile();
			if (excelPacket == null)
				excelPacket = new ExcelPackage(new System.IO.FileInfo(templateFile));
			var sheets = excelPacket.Workbook.Worksheets;
			foreach (var sheet in sheets)
			{
				ExlInterpreter interp = new ExlInterpreter(sheet);
				interp.Complie(this.KeyValues);
            }
            excelPacket.SaveAs(new System.IO.FileInfo(outputFile));
		}
	}
}
