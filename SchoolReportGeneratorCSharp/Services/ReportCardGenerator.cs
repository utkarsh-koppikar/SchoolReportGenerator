using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace SchoolReportGenerator.Services;

/// <summary>
/// Manages the generation of report cards.
/// Uses QuestPDF to generate PDFs matching Word template layout.
/// </summary>
public class ReportCardGenerator
{
    public ReportCardGenerator()
    {
        // Set QuestPDF license (free for open source / personal use)
        QuestPDF.Settings.License = LicenseType.Community;
    }

    /// <summary>
    /// Generate report cards for a given class.
    /// </summary>
    public void GenerateReportCards(string excelPath, string templatePath, string mappingPath, string className, Action<int, int, string>? progressCallback = null)
    {
        Console.WriteLine($"excel_path: {excelPath}");
        Console.WriteLine($"template_path: {templatePath}");
        Console.WriteLine($"mapping_path: {mappingPath}");

        // Setup directories
        var reportCardsDir = $"{className} report_cards";
        EnsureDirectoryExists(reportCardsDir);

        // Process data
        using var dataProcessor = new DataProcessor(excelPath, mappingPath);

        // Get all student rows first to know total count
        var studentRows = dataProcessor.GetStudentRows().ToList();
        var total = studentRows.Count;
        var current = 0;

        // Generate PDF for each student
        foreach (var row in studentRows)
        {
            var studentData = dataProcessor.ProcessStudentData(row, className);
            
            if (studentData == null || !studentData.ContainsKey("name"))
            {
                continue;
            }

            current++;
            var studentName = studentData["name"];
            
            // Report progress
            progressCallback?.Invoke(current, total, studentName);
            Console.WriteLine($"Generating Report card for {studentName} ({current}/{total})");

            var pdfPath = Path.Combine(reportCardsDir, $"{studentName}.pdf");
            CreatePdfReportCard(studentData, pdfPath);
        }
    }

    /// <summary>
    /// Create a PDF report card matching the Word template table layout.
    /// </summary>
    private void CreatePdfReportCard(Dictionary<string, string> studentData, string outputPath)
    {
        Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Size(PageSizes.A4);
                page.Margin(50);
                page.DefaultTextStyle(x => x.FontSize(11));

                page.Content().Column(col =>
                {
                    col.Spacing(0);

                    // Create table matching Word template (GridTable4-Accent5 style)
                    col.Item().Table(table =>
                    {
                        // Define columns (matching Word: 2547 + 4394 ratio ≈ 37% + 63%)
                        table.ColumnsDefinition(columns =>
                        {
                            columns.RelativeColumn(37);
                            columns.RelativeColumn(63);
                        });

                        // Header row style (dark background, white text)
                        var headerStyle = TextStyle.Default.Bold().FontColor(Colors.White);
                        var headerBg = Colors.Blue.Accent2;
                        
                        // Alternating row colors
                        var oddRowBg = Colors.Blue.Lighten4;
                        var evenRowBg = Colors.White;

                        int rowIndex = 0;
                        foreach (var kvp in studentData)
                        {
                            if (kvp.Key == "class" && studentData.ContainsKey("class"))
                            {
                                // Skip if we're iterating and will handle class separately
                            }
                            
                            var isFirstRow = rowIndex == 0;
                            var bgColor = isFirstRow ? headerBg : (rowIndex % 2 == 1 ? oddRowBg : evenRowBg);
                            var textStyle = isFirstRow ? headerStyle : TextStyle.Default;
                            
                            // Label cell (left column)
                            table.Cell()
                                .Background(isFirstRow ? headerBg : Colors.Blue.Lighten3)
                                .Border(1)
                                .BorderColor(Colors.Grey.Lighten1)
                                .Padding(8)
                                .Text(FormatFieldName(kvp.Key))
                                .Bold();

                            // Value cell (right column)
                            table.Cell()
                                .Background(bgColor)
                                .Border(1)
                                .BorderColor(Colors.Grey.Lighten1)
                                .Padding(8)
                                .Text(kvp.Value)
                                .Style(isFirstRow ? headerStyle : TextStyle.Default);

                            rowIndex++;
                        }
                    });
                });
            });
        }).GeneratePdf(outputPath);
    }

    /// <summary>
    /// Format field name for display (capitalize first letter).
    /// </summary>
    private string FormatFieldName(string fieldName)
    {
        if (string.IsNullOrEmpty(fieldName)) return fieldName;
        return char.ToUpper(fieldName[0]) + fieldName.Substring(1);
    }

    /// <summary>
    /// Create directory if it doesn't exist.
    /// </summary>
    private void EnsureDirectoryExists(string directory)
    {
        if (!Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }
}
