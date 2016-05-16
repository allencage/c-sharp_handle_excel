using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;

namespace HandlingExcel
{
	public class ExcelGenerator
	{
		#region Variables
		private string FileName { get; set; }
		private string NewFilePath { get; set; }

		Application xlsxFile = new Application();
		private string[] firstLine = { "Company", "Id", "DateTime"};

		#endregion

		/// <summary>
		/// Contructor
		/// </summary>
		/// <param name="newFilePath"></param>
		/// <param name="fileName"></param>
		public ExcelGenerator(string newFilePath, string fileName)
		{
			NewFilePath = newFilePath;
			FileName = fileName;
		}

		/// <summary>
		/// Simple method for demo purposes, it takes firstLine and outputs every item in the first row
		/// </summary>
		public void WriteFile()
		{
			if (xlsxFile == null)
			{
				Console.WriteLine("Excel is not properly installed ");
				return;
			}

			Workbook xlsxWorkBookW;
			Worksheet xlsxWorkSheet;

			object misValue = System.Reflection.Missing.Value; // when invoking object that need to have a value you can use misvalue not to assign anything to the contructor
															   //read more about missing values here https://msdn.microsoft.com/en-us/library/system.reflection.missing.aspx
			var xlsxFileWorkBooks = xlsxFile.Workbooks; // never use 2 dot syntax with com interop objects without assigning them
			xlsxWorkBookW = xlsxFileWorkBooks.Add(misValue);

			var xlsxWorkBookWorkSheets = xlsxWorkBookW.Worksheets;
			xlsxWorkSheet = xlsxWorkBookWorkSheets.get_Item(1); // gets the first worksheet of the excel

			GenerateFirstExcelLine(xlsxWorkSheet); // this is actually the only line that does something, and it writes the items in the array first line as first row in the excel

			xlsxWorkBookW.SaveCopyAs(NewFilePath.Trim() + FileName.Trim() + ".xlsx"); // after saving we need to close the workbook, be sure to save the file where you have access because else you will get an error
			xlsxWorkBookW.Close(false, misValue, misValue); // syntax for closing the workbook without saving again
			xlsxFile.Quit(); // if you only use this Quit without releasing the com objects, the process will persiste in process manager

			ReleaseObject(xlsxWorkBookW); // release all com objects that were created and assigned to variables
			ReleaseObject(xlsxFileWorkBooks);

			ReleaseObject(xlsxWorkSheet);
			ReleaseObject(xlsxWorkBookWorkSheets);
			ReleaseObject(xlsxFile);
		}

		public void ReleaseObject(object obj)
		{
			try
			{
				Marshal.FinalReleaseComObject(obj); // used to properly discard the com interog objects
				obj = null;
			}
			catch (Exception ex)
			{
				obj = null;
				Console.WriteLine(ex.Message);
			}
			finally
			{
				GC.Collect();
				GC.WaitForPendingFinalizers();
			}
		}

		/// <summary>
		/// The method that actually does something in this demo project
		/// </summary>
		/// <param name="xlWorkSheet"></param>
		public void GenerateFirstExcelLine(Worksheet xlWorkSheet)
		{
			for (int i = 0; i < firstLine.Length; i++)
			{
				xlWorkSheet.Cells[1, i + 1] = firstLine[i];
			}
		}
	}
}