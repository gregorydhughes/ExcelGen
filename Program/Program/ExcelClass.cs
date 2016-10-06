using System.Windows;
using nmspcExceptionClass;
using System;

namespace nmspcExcelOutputClass
{
	class ExcelClass
    {
		private Vector rows;
		private Vector cols;

		private string fileName;

        public ExcelClass()
        {
										
        }

		public ExcelClass(Vector r, Vector c)
		{
			rows = r;
			cols = c;
		}

		public void SetFileName(string fn)
		{
			fileName = fn;
		}
		
		public void Build()
		{

			if (fileName != null)
			{

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