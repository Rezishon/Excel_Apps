using System;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Export.ToDataTable;

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
                            // Assuming the category column is in the second column (change as needed)
                            string category = worksheet.Cells[row, 8].Text;
                            pairs.TryAdd(category, 2);

                            // Create or get a worksheet for the category
                            ExcelWorksheet categoryWorksheet =
                                combinedData.Workbook.Worksheets[category]
                                ?? combinedData.Workbook.Worksheets.Add(category);

                            for (int col = 1; col <= columnCount; col++)
                            {
                                categoryWorksheet.Cells[1, col].Value = worksheet
                                    .Cells[1, col]
                                    .Value;
                            }
                            // Copy the row data to the category worksheet
                            for (int col = 1; col <= columnCount; col++)
                            {
                                try
                                {
                                    categoryWorksheet.Cells[pairs[category], col].Value = worksheet
                                        .Cells[row, col]
                                        .Value;
                                }
                                catch (System.Exception e)
                                {
                                    System.Console.WriteLine(e.Message);
                                }
                            }
                            pairs[category]++;
                        }
                    }
                }

                // Save the combined data to a new Excel file
                string outputFilePath = Path.Combine(directoryPath, "categorized_data.xlsx");
                combinedData.SaveAs(new FileInfo(outputFilePath));

                Console.Clear();
                Console.WriteLine(
                    $"Categorized data saved to {outputFilePath}\nPress any key to exit"
                );
                Console.ReadKey();
            }
            catch (System.Exception e)
            {
                System.Console.WriteLine(e);
                ;
            }
        }
    }
}
