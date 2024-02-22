using System;
using System.IO;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Export.ToDataTable;
using Spectre.Console;

namespace ExcelFileCategorization
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                AnsiConsole.Write(
                    new FigletText("Excel Categorizer").Centered().Color(Color.Purple)
                );

                var rule = new Rule(
                    "[italic blue]Following files are appended to categorized Excel file:[/]"
                );
                rule.LeftJustified();
                AnsiConsole.Write(rule);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                string directoryPath = @"..\..\"; // Specify your directory path
                string fileExtension = ".xlsx"; // Specify the file extension

                ExcelPackage combinedData = new ExcelPackage();

                // Get all Excel files in the specified directory
                string[] excelFiles = Directory.GetFiles(directoryPath, $"*{fileExtension}");
                // Array.Sort(excelFiles);

                Dictionary<string, int> pairs = new Dictionary<string, int>();

                foreach (var excelFile in excelFiles)
                {
                    // AnsiConsole.MarkupLine($"[bold]{Regex.Match(excelFile, @"\w*\W*.xls.$")}[/]");
                    AnsiConsole
                        .Progress()
                        .AutoRefresh(true)
                        .AutoClear(false) // Do not remove the task list when done
                        .HideCompleted(false) // Hide tasks as they are completed
                        .Columns(
                            new ProgressColumn[]
                            {
                                new TaskDescriptionColumn(), // Task description
                                new ProgressBarColumn(), // Progress bar
                                new PercentageColumn(), // Percentage
                                new RemainingTimeColumn(), // Remaining time
                                new SpinnerColumn(), // Spinner
                            }
                        )
                        .Start(ctx =>
                        {
                            // Define tasks
                            var task1 = ctx.AddTask(
                                $"[bold]{Regex.Match(excelFile, @"\w*\W*.xls.$")}[/]"
                            );

                            while (!ctx.IsFinished)
                            {
                                task1.Increment(25);

                                using (var package = new ExcelPackage(new FileInfo(excelFile)))
                                {
                                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming data is in the first sheet

                                    // Read data from the Excel file
                                    int rowCount = worksheet.Dimension.Rows;
                                    int columnCount = worksheet.Dimension.Columns;

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
                                                categoryWorksheet
                                                    .Cells[pairs[category], col]
                                                    .Value = worksheet.Cells[row, col].Value;
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
                        });
                }

                // Save the combined data to a new Excel file
                string outputFilePath = Path.Combine(directoryPath, "categorized_data.xlsx");
                combinedData.SaveAs(new FileInfo(outputFilePath));

                Console.Clear();
                Console.WriteLine(
                    $"Categorized data saved to {outputFilePath}\nPress any key to exit"
                );
                Console.ReadKey();
                Console.Beep();
            }
            catch (System.Exception e)
            {
                System.Console.WriteLine(e);
            }
        }
    }
}
