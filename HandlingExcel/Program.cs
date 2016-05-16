using System.Configuration;

namespace HandlingExcel
{
	class Program
	{
		static void Main(string[] args)
		{
			//For this demo I put the FilePath and FileName in tha App.config
			var excelGenerator = new ExcelGenerator(ConfigurationManager.AppSettings["FilePath"], ConfigurationManager.AppSettings["FileName"]);
			excelGenerator.WriteFile();
		}
	}
}
