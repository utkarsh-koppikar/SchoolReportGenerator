using Avalonia;
using System;
using SchoolReportGenerator.Services;

namespace SchoolReportGenerator;

class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        // If command line args provided, run in headless mode
        if (args.Length >= 4)
        {
            var excelPath = args[0];
            var templatePath = args[1];
            var mappingPath = args[2];
            var className = args[3];

            var generator = new ReportCardGenerator();
            generator.GenerateReportCards(excelPath, templatePath, mappingPath, className);
            Console.WriteLine("Done!");
            return;
        }

        // Otherwise, launch GUI
        BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
    }

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();
}
