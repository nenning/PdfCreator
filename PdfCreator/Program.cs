using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;

using Microsoft.Office.Core;

/// <summary>
/// Add a send to shortcut:
/// cmd -> shell:sendto
/// </summary>
class Program
{
    private static bool hasErrors = false;
    static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            ExitWithError("No arguments. Drag & drop a file or folder.");
        }

        try
        {
            Process(args);
        }
        catch (Exception ex)
        {
            ExitWithError(ex.ToString());
        }

        if (hasErrors)
        {
            ExitWithError("The conversion encountered some issues. See above.");
        }
    }

    private static void ExitWithError(string error)
    {
        var color = Console.ForegroundColor;
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine(error);
        Console.ForegroundColor = color;

        Console.WriteLine();
        Console.WriteLine("Press any key to exit.");
        Console.ReadKey();
        Environment.Exit(0);
    }
    private static void Process(string[] args)
    {
        foreach (string path in args)
        {
            if (Directory.Exists(path))
            {
                ProcessDirectory(path);
            }
            else if (File.Exists(path))
            {
                ProcessFile(path);
            }
            else
            {
                hasErrors = true;
                Console.WriteLine($"Invalid path: {path}");
            }
        }
    }

    static void ProcessDirectory(string directoryPath)
    {
        string[] extensions = [".docx", ".doc", ".ppt", ".pptx", ".xls", ".xlsx"];
        foreach (string subPath in extensions.SelectMany(ext => Directory.EnumerateFiles(directoryPath, "*" + ext, SearchOption.TopDirectoryOnly))) 
        {
            ProcessFile(subPath);
        }
    }

    static void ProcessFile(string filePath)
    {
        Console.WriteLine(filePath);

        string extension = Path.GetExtension(filePath).ToLower();
        string pdfPath = Path.ChangeExtension(filePath, ".pdf");
        if (File.Exists(pdfPath)) return;

        Console.WriteLine("Generating pdf...");
        switch (extension)
        {
            case ".doc":
            case ".docx":
                ConvertWordToPdf(filePath, pdfPath);
                break;
            case ".xls":
            case ".xlsx":
                ConvertExcelToPdf(filePath, pdfPath);
                break;
            case ".ppt":
            case ".pptx":
                ConvertPowerPointToPdf(filePath, pdfPath);
                break;
            default:
                Console.WriteLine($"Unsupported file type: {filePath}");
                hasErrors = true;
                break;
        }
    }

    static void ConvertWordToPdf(string sourcePath, string destinationPath)
    {
        var wordApplication = new Microsoft.Office.Interop.Word.Application();
        try
        {
            var document = wordApplication.Documents.Open(sourcePath, ConfirmConversions: false, ReadOnly: true, AddToRecentFiles: false, Visible: false);
            document.ExportAsFixedFormat(destinationPath, WdExportFormat.wdExportFormatPDF);
            document.Close();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error converting Word file: {ex.Message}");
            hasErrors = true;
        }
        finally
        {
            wordApplication.Quit();
        }
    }

    static void ConvertExcelToPdf(string sourcePath, string destinationPath)
    {
        var excelApplication = new Microsoft.Office.Interop.Excel.Application();
        try
        {
            var workbook = excelApplication.Workbooks.Open(sourcePath, UpdateLinks:false, ReadOnly:true, AddToMru:false);
            workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, destinationPath);
            workbook.Close();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error converting Excel file: {ex.Message}");
            hasErrors = true;
        }
        finally
        {
            excelApplication.Quit();
        }
    }

    static void ConvertPowerPointToPdf(string sourcePath, string destinationPath)
    {
        var powerPointApplication = new Microsoft.Office.Interop.PowerPoint.Application();
        try
        {
            var presentation = powerPointApplication.Presentations.Open(sourcePath, ReadOnly:MsoTriState.msoTrue, Untitled:MsoTriState.msoFalse, WithWindow:MsoTriState.msoFalse);
            presentation.ExportAsFixedFormat(destinationPath, PpFixedFormatType.ppFixedFormatTypePDF, PpFixedFormatIntent.ppFixedFormatIntentPrint);
            presentation.Close();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error converting PowerPoint file: {ex.Message}");
            hasErrors = true;
        }
        finally
        {
            powerPointApplication.Quit();
        }
    }
}
