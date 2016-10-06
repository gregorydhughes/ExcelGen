using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nmspcExcelOutputClass;
using System.IO;

namespace Source
{
    class Program
    {
		static private List<string> filesToGen;

		static private List<string> columnsInXl;

        static void Main(string[] args)
        {
			filesToGen = new List<string>();

			columnsInXl = new List<string>();

			GetNamesOfFilesToGenerate();

			Console.Clear();

			bool unique = SetNamesOfColumnsToGenerateReturnUnique();

			Console.Clear();

			if (!unique)
			{
				Console.WriteLine("Generating non unique excel workbooks...");

				string cat;
				int loc = 1;

				foreach (string file in filesToGen)
				{
					cat = (loc < 10) ? "0" + loc.ToString() : loc.ToString();

					GenerateExcelSheet(cat+file, columnsInXl);

					loc++;
				}
			}
			else
			{
				Console.WriteLine("Not generating unique excel workbooks...");
			}
        }

		static void GetNamesOfFilesToGenerate()
		{
			Console.WriteLine("Please enter the name of the file with the names of the excel sheets you would like to generate:");

			string fName = Console.ReadLine();

			while (!File.Exists(fName))
			{
				Console.WriteLine("File not found, please re-enter file:");
				fName = Console.ReadLine();
			}

			ReadFileIntoList(fName, ref filesToGen);
		}

		static bool SetNamesOfColumnsToGenerateReturnUnique()
		{

			if (GetAnswerToQuestion("Do you want to generate the same columns for all sheets?"))
			{
				Console.WriteLine("Please enter the name of the file with the names of the columns you would like to insert into all the sheets:");

				string fName = Console.ReadLine();

				while (!File.Exists(fName))
				{
					Console.WriteLine("File not found, please re-enter file:");
					fName = Console.ReadLine();
				}

				ReadFileIntoList(fName, ref columnsInXl);

				return false;
			}
			else
			{
				// oh bother
				// iterate through each excel sheet and ask the names of the columns to generate for each sheet

				return true;
			}
		}

		static void ReadFileIntoList(string fName, ref List<string> l)
		{
			StreamReader sR = new StreamReader(fName);

			string line = sR.ReadLine();

			while (line != null)
			{
				l.Add(line);

				line = sR.ReadLine();
			}

			sR.Close();
		}

		static void GenerateExcelSheet(string file, List<string> cols)
		{
			ExcelClass ec = new ExcelClass(cols);

			ec.fileName = file;

			ec.workSheetName = "Music";

			ec.Build();
		}

		static bool GetAnswerToQuestion(string question)
		{
			Console.WriteLine(question);

			char ans = Convert.ToChar(Console.ReadLine().ToLower()[0]);

			switch (ans)
			{
				case 'y':
					return true;
					break;
			}

			return false;
		}
    }
}
