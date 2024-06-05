using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;


//确保安装了 EPPlus 或 ClosedXML 库来读取 Excel 文件。在项目中，可以通过 NuGet 包管理器安装这些库。
class Program
{
    static void Main()
    {
        string filePath = "d:/煤样数据.xlsx"; // Excel file path
        Dictionary<string, List<double>> data = ReadExcelFile(filePath);

        if (data != null && data.Count > 0)
        {
            foreach (var entry in data)
            {
                string indicator = entry.Key;
                List<double> values = entry.Value;

                CalculateStatistics(values, out double max, out double min, out double sum, out double avg);
                Console.WriteLine($"Indicator: {indicator} - Max: {max}, Min: {min}, Sum: {sum}, Avg: {avg}");
            }
        }
        else
        {
            Console.WriteLine("No data found in the Excel file.");
        }
        Console.ReadKey();
    }

    static Dictionary<string, List<double>> ReadExcelFile(string filePath)
    {
        var data = new Dictionary<string, List<double>>();
        FileInfo fileInfo = new FileInfo(filePath);

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Add this line

        using (var package = new ExcelPackage(fileInfo))
        {
            if (package.Workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("No worksheets found in the Excel file.");
                return data;
            }

            ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming data is in the first sheet
            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            // Read header
            var headers = new List<string>();
            for (int col = 2; col <= colCount; col++) // Start from the second column
            {
                headers.Add(worksheet.Cells[1, col].Text);
                data[worksheet.Cells[1, col].Text] = new List<double>();
            }

            // Read data
            for (int row = 2; row <= rowCount; row++) // Assuming first row is header
            {
                for (int col = 2; col <= colCount; col++) // Start from the second column
                {
                    if (double.TryParse(worksheet.Cells[row, col].Text, out double value))
                    {
                        data[headers[col - 2]].Add(value); // Adjust index to match headers list
                    }
                }
            }
        }

        return data;
    }

    static void CalculateStatistics(List<double> data, out double max, out double min, out double sum, out double avg)
    {
        max = double.MinValue;
        min = double.MaxValue;
        sum = 0;

        foreach (var value in data)
        {
            if (value > max)
                max = value;
            if (value < min)
                min = value;
            sum += value;
        }

        avg = data.Count > 0 ? sum / data.Count : 0;
    }
}

