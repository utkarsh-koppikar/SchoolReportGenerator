using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using System;
using System.Threading.Tasks;
using SchoolReportGenerator.Services;

namespace SchoolReportGenerator;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        // Hide progress section initially
        ProgressSection.IsVisible = false;
    }

    private async void BrowseExcel_Click(object? sender, RoutedEventArgs e)
    {
        var path = await PickFile("Excel Files", "xlsx");
        if (path != null) ExcelPathBox.Text = path;
    }

    private async void BrowseWord_Click(object? sender, RoutedEventArgs e)
    {
        var path = await PickFile("Word Files", "docx");
        if (path != null) WordPathBox.Text = path;
    }

    private async void BrowseMapping_Click(object? sender, RoutedEventArgs e)
    {
        var path = await PickFile("JSON Files", "json");
        if (path != null) MappingPathBox.Text = path;
    }

    private async Task<string?> PickFile(string filterName, string extension)
    {
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = $"Select {filterName}",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType(filterName) { Patterns = new[] { $"*.{extension}" } }
            }
        });

        return files.Count > 0 ? files[0].Path.LocalPath : null;
    }

    private async void Generate_Click(object? sender, RoutedEventArgs e)
    {
        var excelPath = ExcelPathBox.Text;
        var wordPath = WordPathBox.Text;
        var mappingPath = MappingPathBox.Text;
        var className = ClassNameBox.Text;

        if (string.IsNullOrEmpty(excelPath) || string.IsNullOrEmpty(wordPath) ||
            string.IsNullOrEmpty(mappingPath) || string.IsNullOrEmpty(className))
        {
            ShowStatus("Please select all files and enter class name", isError: true);
            return;
        }

        try
        {
            // Disable button and show progress
            GenerateButton.IsEnabled = false;
            ProgressSection.IsVisible = true;
            StatusBorder.IsVisible = false;
            
            // Reset progress
            ProgressBar.Value = 0;
            ProgressCountText.Text = "Starting...";
            CurrentStudentText.Text = "";
            RemainingText.Text = "";

            var generator = new ReportCardGenerator();
            
            // Progress callback
            Action<int, int, string> progressCallback = (current, total, studentName) =>
            {
                Dispatcher.UIThread.Post(() =>
                {
                    var percent = (double)current / total * 100;
                    var remaining = total - current;
                    
                    ProgressBar.Value = percent;
                    ProgressCountText.Text = $"{current}/{total} done";
                    CurrentStudentText.Text = studentName;
                    RemainingText.Text = remaining > 0 ? $"{remaining} remaining" : "Almost done...";
                });
            };

            await Task.Run(() => generator.GenerateReportCards(excelPath, wordPath, mappingPath, className, progressCallback));
            
            // Complete
            ProgressBar.Value = 100;
            ProgressCountText.Text = "All done!";
            CurrentStudentText.Text = "";
            RemainingText.Text = "";
            
            ShowStatus($"✓ Report cards generated successfully in '{className} report_cards' folder!", isError: false);
        }
        catch (Exception ex)
        {
            ShowStatus($"✗ Error: {ex.Message}", isError: true);
        }
        finally
        {
            GenerateButton.IsEnabled = true;
        }
    }

    private void ShowStatus(string message, bool isError)
    {
        StatusBorder.IsVisible = true;
        StatusBorder.Background = isError 
            ? new SolidColorBrush(Color.FromRgb(255, 200, 200))  // Light red
            : new SolidColorBrush(Color.FromRgb(200, 255, 200)); // Light green
        StatusText.Text = message;
        StatusText.Foreground = isError 
            ? new SolidColorBrush(Color.FromRgb(180, 0, 0))      // Dark red
            : new SolidColorBrush(Color.FromRgb(0, 128, 0));     // Dark green
    }
}
