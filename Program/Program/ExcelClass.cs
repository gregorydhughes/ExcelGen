using System.Windows;
using System.Collections.Generic;
using System.Text;
using nmspcExceptionClass;
using Excel = Microsoft.Office.Interop.Excel;
using System;

namespace nmspcExcelOutputClass
{
	class ExcelClass
    {
		private const string DEFAULT_WORKSHEET_NAME = "Sheet1";

		private List<string> rows;
		private List<string> cols;

		private string fn;

		public string fileName
		{
			get
			{
				return fn;
			}
			set
			{
				fn = value;
			}
		}

		private string wsn;
		
		public string workSheetName
		{
			get
			{
				return wsn;
			}
			set
			{
				wsn = value;
			}
		}

        public ExcelClass()
        {
			workSheetName = DEFAULT_WORKSHEET_NAME;
		}

		public ExcelClass(List<string> c)
		{
			cols = c;
			workSheetName = DEFAULT_WORKSHEET_NAME;
		}

		public ExcelClass(List<string> r, List<string> c)
		{
			rows = r;
			cols = c;
			workSheetName = DEFAULT_WORKSHEET_NAME;
		}
		
		public void Build()
		{

			if (fileName != null)
			{
				Excel.Application xlApp = new Excel.Application();

				xlApp.Visible = false;

				Excel.Workbook wb = xlApp.Workbooks.Add();

				Excel.Worksheet ws = wb.Sheets.Add();

				ws.Name = workSheetName;

				int pos = 1;

				foreach (string col in cols)
				{
					ws.Cells[1, pos] = col;
					pos++;
				}

				wb.SaveAs(@"D:\Users\Gregory\Documents\MusicSheets\" + fileName);

				wb.Close();

				xlApp.Quit();
			}
			else
			{
				//throw new NoFileNameGivenException();
				Console.WriteLine("No file name was given...\n\nWe're going to abort...");
				Environment.Exit(-1);
			}
		} 
    }
}