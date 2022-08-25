using System.Drawing;
using Create_Excel_XLSX.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Create_Excel_XLSX;

public class ExportOperation
{
    public async Task GenerateExcel<T>(IEnumerable<T> data)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var file = new FileInfo(@"C:\your\root\name\ExcelDemo.xlsx");

        await SaveExcelFile(data, file);

        var personFromExcel = await LoadExcelFile(file);

        personFromExcel.ForEach(p => { Console.WriteLine($"{p.Id} {p.FirstName} {p.LastName}"); });
    }

    private static async Task SaveExcelFile<T>(IEnumerable<T> data, FileInfo file)
    {
        DeleteIfExists(file);

        using var package = new ExcelPackage(file);

        var ws = package.Workbook.Worksheets.Add("Report");

        // Formats the header 
        ws.Cells["A1"].Value = "Header";
        ws.Cells["A1:F1"].Merge = true;

        ws.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        ws.Row(1).Style.Font.Size = 24;
        ws.Row(1).Style.Font.Color.SetColor(Color.Blue);

        ws.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        ws.Row(2).Style.Font.Bold = true;
        ws.Column(3).Width = 20;

        var range = ws.Cells["A2"].LoadFromCollection(data, true);

        // Formats the columns
        range.Style.Font.Size = 12;
        range.Style.Font.Color.SetColor(Color.Black);
        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        range.AutoFitColumns();

        await package.SaveAsync();
    }

    private static void DeleteIfExists(FileSystemInfo file)
    {
        if (file.Exists) file.Delete();
    }

    private async Task<List<User>> LoadExcelFile(FileInfo file)
    {
        List<User> output = new();

        using var package = new ExcelPackage(file);

        await package.LoadAsync(file);

        var ws = package.Workbook.Worksheets[0];

        var row = 3;

        const int col = 1;

        while (string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()) == false)
        {
            User p = new()
            {
                Id = int.Parse(ws.Cells[row, col].Value.ToString()!),
                FirstName = ws.Cells[row, col + 1].Value.ToString()!,
                LastName = ws.Cells[row, col + 2].Value.ToString()!
            };

            output.Add(p);
            row += 1;
        }

        return output;
    }
}