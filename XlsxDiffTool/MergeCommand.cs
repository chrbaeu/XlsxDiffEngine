using CliFx;
using CliFx.Binding;
using CliFx.Infrastructure;
using System.IO;
using XlsxDiffEngine;

namespace XlsxDiffTool;

[Command("merge", Description = "Merges Excel files and saves the merged data in an output file.")]
/// <remarks>
/// Examples:
/// xlsxdifftool merge --sourcePath "C:\Data\ExcelFiles" --mergeResultPath "MergedData.xlsx" --fromRow 7
/// xlsxdifftool merge --sourcePath "Old.xlsx" "New.xlsx" --mergeResultPath "MergedData.xlsx" --worksheetNameColumnName "Worksheet" --documentNameColumnName "Document"
/// </remarks>
internal sealed partial class MergeCommand : ICommand
{
    [CommandOption("sourcePath", 's', Description = "Path for a source .xlsx file or a folder containing .xlsx files.")]
    public required string[] SourcePaths { get; set; } = [];

    [CommandOption("mergeResultPath", 'm', Description = "Path for the resulting merged .xlsx file.")]
    public string MergePath { get; set; } = "MergedData.xlsx";

    [CommandOption("fromRow", 'r', Description = "First row to read from each source worksheet.")]
    public int FromRow { get; set; } = 1;

    [CommandOption("worksheetNameColumnName", Description = "Optional column name for source worksheet names.")]
    public string? WorksheetNameColumnName { get; set; }

    [CommandOption("documentNameColumnName", Description = "Optional column name for source document names.")]
    public string? DocumentNameColumnName { get; set; }

    [CommandOption("mergedDocumentName", Description = "Optional worksheet name for the merged output.")]
    public string? MergedDocumentName { get; set; }

    [CommandOption("worksheetName", 'w', Description = "Optional source worksheet name to include. Can be specified multiple times.")]
    public string[] WorksheetNames { get; set; } = [];

    public ValueTask ExecuteAsync(IConsole console)
    {
        if (SourcePaths.Length == 0)
        {
            throw new InvalidOperationException("At least one source path must be specified.");
        }
        if (FromRow < 1)
        {
            throw new InvalidOperationException("The first row must be greater than zero.");
        }

        List<string> files = GetSourceFiles();
        if (files.Count == 0)
        {
            throw new FileNotFoundException("No source .xlsx files were found.");
        }

        string outputPath = GetOutputPath();
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? ".");

        using XlsxDataProvider xlsxDataProvider = new(files.Select(CreateXlsxFileInfo).ToList(), new XlsxDataProviderConfig()
        {
            MergeDocuments = true,
            MergeWorksheets = true,
            WorksheetNameColumnName = string.IsNullOrWhiteSpace(WorksheetNameColumnName) ? null : WorksheetNameColumnName,
            DocumentNameColumnName = string.IsNullOrWhiteSpace(DocumentNameColumnName) ? null : DocumentNameColumnName,
            MergedDocumentName = string.IsNullOrWhiteSpace(MergedDocumentName) ? null : MergedDocumentName,
            WorksheetNames = WorksheetNames.Length == 0 ? null : WorksheetNames,
        });

        xlsxDataProvider.SaveAs(new FileInfo(outputPath));
        console.Output.WriteLine($"Merged file successfully saved to '{outputPath}'.");
        return default;
    }

    private List<string> GetSourceFiles()
    {
        List<string> files = [];
        foreach (string sourcePath in SourcePaths)
        {
            if (File.Exists(sourcePath))
            {
                if (!string.Equals(Path.GetExtension(sourcePath), ".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException($"The file '{sourcePath}' is not an .xlsx file.");
                }
                files.Add(sourcePath);
            }
            else if (Directory.Exists(sourcePath))
            {
                files.AddRange(Directory.GetFiles(sourcePath, "*.xlsx"));
            }
            else
            {
                throw new FileNotFoundException($"The source path '{sourcePath}' was not found.");
            }
        }

        return files
            .Where(x => !Path.GetFileName(x).StartsWith("~$", StringComparison.Ordinal))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Order(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private string GetOutputPath()
    {
        if (string.IsNullOrWhiteSpace(MergePath))
        {
            return "MergedData.xlsx";
        }
        if (Directory.Exists(MergePath) || string.IsNullOrEmpty(Path.GetExtension(MergePath)))
        {
            return Path.Combine(MergePath, "MergedData.xlsx");
        }
        if (!string.Equals(Path.GetExtension(MergePath), ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("The merge result path must be an .xlsx file.");
        }
        return MergePath;
    }

    private XlsxFileInfo CreateXlsxFileInfo(string path)
        => new(path) { FromRow = FromRow };
}
