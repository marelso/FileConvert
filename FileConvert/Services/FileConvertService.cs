using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentResults;
using OfficeOpenXml;

namespace FileConvert.Services
{
    internal class FileConvertService
    {
        static public Result ConvertToCsv(string source, string sheet = null)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var inputFile = new FileInfo(source);
            try
            {
                using (var doc = new ExcelPackage(inputFile))
                {
                    var workbook = doc.Workbook;
                    ExcelWorksheet worksheet;

                    #region Validanting workbook
                    if (workbook == null)
                        return Result.Fail($"Error tring to read workbook.\n");

                    if (workbook.Worksheets.Count == 0)
                        return Result.Fail($"Workbook doens't contain any worksheet.\n");
                    #endregion

                    #region Setup and validation worksheet
                    if (!string.IsNullOrEmpty(sheet))
                        worksheet = workbook.Worksheets.FirstOrDefault(x => x.Name == sheet);
                    else
                        worksheet = workbook.Worksheets[1];
                    if (worksheet == null)
                        return Result.Fail($"Error: Worksheet is empty.\n");
                    #endregion

                    int columnNumber = worksheet.Dimension.End.Column;
                    var convertedRecords = new List<List<string>>(worksheet.Dimension.End.Row);
                    var excelRows = worksheet.Cells.GroupBy(c => c.Start.Row).ToList();

                    excelRows.ForEach(r =>
                    {
                        var currentRecord = new List<string>(columnNumber);
                        var cells = r.OrderBy(cell => cell.Start.Column).ToList();
                        for (int i = 0; i < columnNumber; i++)
                        {
                            var currentCell = cells.Where(c => c.Start.Column == i).FirstOrDefault();
                            if (currentCell == null)
                                AddCellValue(string.Empty, currentRecord);
                            else
                                AddCellValue(currentCell.Value == null ? string.Empty : currentCell.Value.ToString(), currentRecord);
                        }
                        convertedRecords.Add(currentRecord);
                    }
                    );

                    string outputFile = source.Replace(".xlsx", string.Format("_{0}.csv", worksheet.Name));
                    WriteToFile(convertedRecords, outputFile);

                    return Result.Ok().WithSuccess("File converted with success.\n");
                }
            }
            catch (Exception e)
            {
                return Result.Fail($"Error: {e.Message}\n");
            }
        }
        private static void AddCellValue(string text, List<string> record)
        {
            record.Add(text);
        }
        private static void WriteToFile(List<List<string>> records, string path)
        {
            var commaDelimited = new List<string>(records.Count);
            records.ForEach(r => commaDelimited.Add(r.ToDelimitedString(",")));
            File.WriteAllLines(path, commaDelimited);
        }
    }
    public static class ListExtension
    {
        public static string ToDelimitedString(this List<string> list, string separator = ":", bool insertSpaces = false)
        {
            var result = string.Empty;

            for(int i = 0; i < list.Count; i++)
            {
                var currentString = list[i];
                if( i < (list.Count - 1))
                {
                    currentString += separator;
                    if (insertSpaces)
                        currentString += ' ';
                }
                result += currentString;
            }
            return result;
        }
    }
}
