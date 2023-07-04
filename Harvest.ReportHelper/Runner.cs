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

public class Runner
{
    private string? _currentFilePath;

    public Runner()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }
    
    public async Task<int> GetOptionAsync()
    {
        await Console.Out.WriteLineAsync("");
        await Console.Out.WriteLineAsync("Choose: ");
        await Console.Out.WriteLineAsync("1) Check for issues (Default)");
        await Console.Out.WriteLineAsync("2) Clean");
        await Console.Out.WriteLineAsync("3) Clean & split");
        
        var optionValue = Console.ReadLine();
        
        return Int32.TryParse(optionValue, out var option) ? option : 0;
    }
        
    public async Task<bool> TryGetFileNameAsync()
    {
        await Console.Out.WriteAsync("Enter source file name: ");
        var fileReference = Console.ReadLine();

        if (String.IsNullOrWhiteSpace(fileReference))
            return false;
        
        _currentFilePath = fileReference;
        return true;
    }

    public async Task RunCheckAsync()
    {
        if (_currentFilePath == null)
            return;
        
        var shortenedPath = Path.Combine("%USERPROFILE%\\Downloads", _currentFilePath);
        var filePath = Environment.ExpandEnvironmentVariables(shortenedPath);
        var file = new FileInfo(filePath);

        var codeRefRegex = new Regex(@"^(?<prefix>[A-Z]{2})[A-Z]{1,}-");
        var hoursRegex = new Regex(@"^\d+(\.(25|5|75))?$");
        var rowsWithIssues = new Collection<int>();
        
        using (var package = new ExcelPackage(file))
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

    public async Task RunCleanAsync(bool deleteClientColumn = true, bool allowPrefix = true)
    {
        if (_currentFilePath == null)
            return;
        
        var shortenedPath = Path.Combine("%USERPROFILE%\\Downloads", _currentFilePath);
        var filePath = Environment.ExpandEnvironmentVariables(shortenedPath);

        string? prefix = null;
        if (allowPrefix)
        {
            await Console.Out.WriteAsync("Enter prefix (optional): ");
            prefix = Console.ReadLine();
        }
        
        var file = new FileInfo(filePath);
        
        using (var package = new ExcelPackage(file))
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
                _currentFilePath = filePath.Replace("harvest_", $"{prefix.Trim()}_");
                
                await package.SaveAsAsync(_currentFilePath);
                await Console.Out.WriteLineAsync($"Cleaned and output to {_currentFilePath}");
            }
            else
            {
                await package.SaveAsAsync(filePath);
                await Console.Out.WriteLineAsync($"Cleaned {filePath}");
            }
        }
    }

    public async Task RunSplitAsync()
    {
        if (_currentFilePath == null)
            return;
        
        var shortenedPath = Path.Combine("%USERPROFILE%\\Downloads", _currentFilePath);
        var filePath = Environment.ExpandEnvironmentVariables(shortenedPath);
        var file = new FileInfo(filePath);
        var clientNames = new List<string>();
        
        using (var package = new ExcelPackage(file))
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
        if (_currentFilePath == null)
            return;
        
        var shortenedPath = Path.Combine("%USERPROFILE%\\Downloads", _currentFilePath);
        var filePath = Environment.ExpandEnvironmentVariables(shortenedPath);
        var file = new FileInfo(filePath);

        using (var package = new ExcelPackage(file))
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

            var splitFilePath = filePath.Replace("harvest_", $"{clientName.ToLower().Trim()}_");
            
            await package.SaveAsAsync(splitFilePath);
            await Console.Out.WriteLineAsync($"Split and output to {splitFilePath}");
        }
    }
}