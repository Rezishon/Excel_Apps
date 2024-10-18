using System.Text;
using ClosedXML.Excel;
namespace Excel_To_SqlQ;

class Program
{
    static void Main(string[] args)
    {
        // These should be changed in future: - Excel file path - Worksheet number - Table name -
        using var workbook = new XLWorkbook("./all1.xlsx");
        // Access the worksheet you want to read from
        var worksheet = workbook.Worksheet(1);

        Processes processes = new Processes(worksheet, "Costumer");

        processes.Luncher();
    }
}

internal class Processes(IXLWorksheet worksheet, string TableName)
{
    public List<string> Headers { get; set; } = new List<string>();
    public IXLWorksheet Worksheet { get; set; } = worksheet;
    public List<string> RowsData { get; set; } = new List<string>();

    public void Luncher()
    {
        HeaderFounder();
        DataFounder();
    }

    public void HeaderFounder()
    {
        foreach (var cell in Worksheet.Row(1).Cells())
        {
            Headers.Add(cell.Value.ToString());
        };
        Worksheet.Row(1).Delete();
    }
    public void DataFounder()
    {
        foreach (var row in worksheet.Rows())
        {
            foreach (var cell in row.Cells())
            {
                RowsData.Add(cell.Value.ToString());
            }

            Combiner();
            RowsData.Clear();

        }
    }

    public void Combiner()
    {
    }
}
