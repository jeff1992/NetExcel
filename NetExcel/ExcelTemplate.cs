using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetExcel
{
	public class ExcelTemplate
    {
		/// <summary>
		/// 
		/// </summary>
		/// <param name="templateFile">template file</param>
		public ExcelTemplate(string templateFile)
		{
			this.templateFile = templateFile;
			CheckFile();
        }


		ExcelPackage excelPacket;
		string templateFile;

		/// <summary>
		/// Get the key value dictionary for render work
		/// </summary>
		public Dictionary<string, object> Values { get; } = new Dictionary<string, object>();

		void CheckFile()
		{
			if (!System.IO.File.Exists(templateFile))
				throw new Exception($"File \"{templateFile}\" does not exist");
		}
        /// <summary>
        /// Save file
        /// </summary>
        /// <param name="fileName"></param>
        public void SaveAs(string fileName)
        {
            this.SaveAs(fileName, null);
        }
        /// <summary>
        /// Save and set file with password to modify
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="password"></param>
        public void SaveAs(string fileName, string password)
        {
            CheckFile();
            var package = new ExcelPackage(new System.IO.FileInfo(templateFile));
            foreach (var sheet in package.Workbook.Worksheets)
            {
                ExlInterpreter interp = new ExlInterpreter(sheet);
                interp.Complie(this.Values);
                if (!string.IsNullOrWhiteSpace(password))
                    sheet.Protection.SetPassword(password);
            }
            package.SaveAs(new System.IO.FileInfo(fileName));
        }
	}
}
