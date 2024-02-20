using System;
using System.IO;
using OfficeOpenXml;

namespace ExcelFileCategorization
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.Write("Working...");
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                string directoryPath = @"..\..\"; // Specify your directory path
                string fileExtension = ".xlsx"; // Specify the file extension

                ExcelPackage combinedData = new ExcelPackage();

                // Get all Excel files in the specified directory
                string[] excelFiles = Directory.GetFiles(directoryPath, $"*{fileExtension}");

                foreach (var excelFile in excelFiles)
                {
                    using (var package = new ExcelPackage(new FileInfo(excelFile)))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming data is in the first sheet

                        // Read data from the Excel file
                        int rowCount = worksheet.Dimension.Rows;
                        int columnCount = worksheet.Dimension.Columns;

                        Dictionary<string, int> pairs = new Dictionary<string, int>();

                        for (int row = 2; row <= rowCount; row++) // Assuming the first row is the header
                        {
                        }
                }
            }
            catch (System.Exception e)
            {
                System.Console.WriteLine(e);
                ;
            }
        }
    }
}
