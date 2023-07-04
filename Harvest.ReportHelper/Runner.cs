using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Harvest.ReportHelper;

public partial class Runner
{
    [GeneratedRegex("^(?<prefix>[A-Z]{2})[A-Z]{1,}-")]
    private static partial Regex CodeRefRegex();

    [GeneratedRegex("^\\d+(\\.(25|5|75))?$")]
    private static partial Regex HoursRegex();

    private FileInfo? _currentFileInfo;

    public Runner()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }
    
    public static async Task<int> GetOptionAsync()
    {
        await Console.Out.WriteLineAsync("");
        await Console.Out.WriteLineAsync("Choose: ");
        await Console.Out.WriteLineAsync("1) Check for issues");
        await Console.Out.WriteLineAsync("2) Clean");
        await Console.Out.WriteLineAsync("3) Clean & split");
        
        var optionValue = Console.ReadLine();
        
        return Int32.TryParse(optionValue, out var option) ? option : 0;
    }
        
    public async Task<bool> TryGetFileNameAsync()
    {
        if (_currentFileInfo != null)
        {
            await Console.Out.WriteLineAsync("Use current file? [Y]/n");
            if ((Console.ReadLine()?.ToUpper() ?? "Y") != "N")
            {
                return true;
            }
        }
        
        await Console.Out.WriteAsync("Enter source file name: ");
        var fileReference = Console.ReadLine();

        if (String.IsNullOrWhiteSpace(fileReference))
            return false;
        
        FileInfo? file;
        
        if (fileReference.Contains(@":\"))
        {
            file = new FileInfo(fileReference.Trim('"'));
        }
        else
        {
            var shortenedPath = Path.Combine("%USERPROFILE%\\Downloads", fileReference);
            file = new FileInfo(Environment.ExpandEnvironmentVariables(shortenedPath));
        }

        if (!file.Exists)
        {
            await Console.Out.WriteLineAsync("File doesn't exist");
            return false;
        }

        _currentFileInfo = file;
        return true;
    }

    public async Task RunCheckAsync()
    {
        try
        {
            var codeRefRegex = CodeRefRegex();
            var hoursRegex = HoursRegex();
            
            var rowsWithIssues = new Collection<int>();
            
            using (var package = new ExcelPackage(_currentFileInfo))
            {
                var sheet = package.Workbook.Worksheets[0];

                const int projectCodeColumn = 4;
                const int notesColumn = 6;
                const int hoursColumn = 7;

                for (var row = 2; row < sheet.Rows.EndRow; row++)
                {
                    var projectCodeValue = sheet.Cells[row, projectCodeColumn].Value?.ToString();
                    if (projectCodeValue == null)
                        break;
                    
                    var notesValue = sheet.Cells[row, notesColumn].Value?.ToString();
                    if (!String.IsNullOrWhiteSpace(notesValue) && codeRefRegex.IsMatch(notesValue))
                    {
                        var notesCodePrefix = codeRefRegex.Match(notesValue).Groups["prefix"].Value;
                        if (!projectCodeValue.StartsWith(notesCodePrefix))
                        {
                            rowsWithIssues.Add(row);
                            sheet.Row(row).Style.Font.Bold = true;
                            sheet.Cells[row, notesColumn].Style.Font.Color.SetColor(Color.Crimson);
                            sheet.Cells[row, notesColumn].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[row, notesColumn].Style.Fill.BackgroundColor.SetColor(Color.MistyRose);
                        }
                    }

                    var hoursValue = sheet.Cells[row, hoursColumn].Value?.ToString();
                    if (!String.IsNullOrWhiteSpace(hoursValue) && !hoursRegex.IsMatch(hoursValue))
                    {
                        rowsWithIssues.Add(row);
                        sheet.Row(row).Style.Font.Bold = true;
                        sheet.Cells[row, hoursColumn].Style.Font.Color.SetColor(Color.Crimson);
                        sheet.Cells[row, hoursColumn].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[row, hoursColumn].Style.Fill.BackgroundColor.SetColor(Color.MistyRose);
                    }
                }

                if (rowsWithIssues.Any())
                {
                    sheet.Column(notesColumn).Width = 70;

                    await package.SaveAsync();
                    await Console.Out.WriteLineAsync($"Checked with {rowsWithIssues.Count} issues");
                }
                else
                {
                    await Console.Out.WriteLineAsync($"Checked with no issues");
                }
            }
        }
        catch (Exception ex)
        {
            await Console.Out.WriteLineAsync(ex.Message);
            Console.ReadLine();
        }
    }

    public async Task RunCleanAsync(bool deleteClientColumn = true, bool allowPrefix = true)
    {
        string? prefix = null;
        if (allowPrefix)
        {
            await Console.Out.WriteAsync("Enter prefix (optional): ");
            prefix = Console.ReadLine();
        }
        
        using (var package = new ExcelPackage(_currentFileInfo))
        {
            var sheet = package.Workbook.Worksheets[0];
            if (sheet.Cells[1, 7].Value?.ToString() != "Hours")
            {
                await Console.Out.WriteLineAsync("Unable to clean - Columns do not match expected positions");
                return;
            }
            
            sheet.Column(6).Width = 70; // Notes
            sheet.Column(8).Width = 10; // Rounded Hours
            sheet.Cells[1, 8].Value = "Hours";

            
            var columnsToDelete = new List<int>
            {
                2,  // Client
                7,  // Hours
                9,  // Billable?
                10, // Invoiced?
                13, // Roles
                14, // Employee
                17, // Cost Rate
                18, // Cost amount
                19, // Currency
                20  // External Reference URL
            };
            
            columnsToDelete.Reverse();
            
            foreach (var x in columnsToDelete)
            {
                if (!deleteClientColumn && x == 2)
                    continue;
                
                sheet.DeleteColumn(x);
            }
            
            for (var row = 2; row < sheet.Rows.EndRow; row++)
            {
                var projectCodeValue = sheet.Cells[row, 1].Value?.ToString();
                if (projectCodeValue == null)
                    break;

                if (!sheet.Row(row).Style.Fill.BackgroundColor.Auto || !sheet.Row(row).Style.Font.Color.Auto)
                {
                    sheet.Row(row).Style.Font.Bold = false;
                    sheet.Row(row).Style.Font.Color.SetAuto();
                    sheet.Row(row).Style.Fill.PatternType = ExcelFillStyle.None;
                }
            }

            if (!String.IsNullOrWhiteSpace(prefix))
            {
                var fileName = _currentFileInfo?.Name.Replace("harvest_", $"{prefix.Trim()}_");
                
                await package.SaveAsAsync(fileName);
                await Console.Out.WriteLineAsync($"Cleaned and output to {fileName}");
            }
            else
            {
                await package.SaveAsAsync(_currentFileInfo);
                await Console.Out.WriteLineAsync($"Cleaned {_currentFileInfo}");
            }
        }
    }

    public async Task RunSplitAsync()
    {
        var clientNames = new List<string>();
        
        using (var package = new ExcelPackage(_currentFileInfo))
        {
            var sheet = package.Workbook.Worksheets[0];
            
            for (var row = 2; row < sheet.Rows.EndRow; row++)
            {
                var clientName = sheet.Cells[row, 2].Value?.ToString();
                if (clientName == null)
                    break;
                
                if (!clientNames.Contains(clientName))
                {
                    clientNames.Add(clientName);
                }
            }
        }

        foreach (var clientName in clientNames)
        {
            await SplitForClientAsync(clientName);
        }
    }

    private async Task SplitForClientAsync(string clientName)
    {
        if (_currentFileInfo == null)
            return;
        
        using (var package = new ExcelPackage(_currentFileInfo))
        {
            var sheet = package.Workbook.Worksheets[0];
            var rowsToDelete = new List<int>();
            
            for (var row = 2; row < sheet.Rows.EndRow; row++)
            {
                var columnValue = sheet.Cells[row, 2].Value?.ToString();
                if (columnValue == null)
                    break;
                
                if (columnValue == clientName)
                    continue;

                rowsToDelete.Add(row);
            }

            rowsToDelete.Reverse();
            
            sheet.DeleteColumn(2);
            foreach (var row in rowsToDelete)
            {
                sheet.DeleteRow(row);   
            }

            var splitFilePath = _currentFileInfo.Name.Replace("harvest_", $"{clientName.ToLower().Trim().Replace(" ", "_")}_");
            
            await package.SaveAsAsync(splitFilePath);
            await Console.Out.WriteLineAsync($"Split and output to {splitFilePath}");
        }
    }
}