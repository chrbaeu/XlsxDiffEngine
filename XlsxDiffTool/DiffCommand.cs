using CliFx;
using CliFx.Attributes;
using CliFx.Infrastructure;
using XlsxDiffTool.Models;
using XlsxDiffTool.Services;
using System.IO;

namespace XlsxDiffTool;

[Command("diff", Description = "Compares two Excel files and saves the differences in an output file.")]
internal sealed class DiffCommand(
    DiffConfigModel diffConfigModel,
    DiffConfigService diffConfigService,
    ExcelDiffService excelDiffService
    ) : ICommand
{

    [CommandOption("oldPath", 'o', IsRequired = true, Description = "Path for the older src xlsx data.")]
    public string OldPath { get; init; } = "";

    [CommandOption("newPath", 'n', IsRequired = true, Description = "Path for the newer src xlsx data.")]
    public string NewPath { get; init; } = "";

    [CommandOption("configPath", 'c', Description = "Path for the configuration to be used for the diff process.")]
    public string ConfigPath { get; init; } = "";

    [CommandOption("diffResultPath", 'd', Description = "Path for the resutlig diff xlsx file.")]
    public string DiffPath { get; init; } = ".";

    public async ValueTask ExecuteAsync(IConsole console)
    {
        if (!File.Exists(OldPath))
        {
            throw new FileNotFoundException($"The file {OldPath} was not found.”");
        }
        if (!File.Exists(NewPath))
        {
            throw new FileNotFoundException($"The file {NewPath} was not found.");
        }
        if (string.IsNullOrEmpty(ConfigPath))
        {
            diffConfigService.Reset();
        }
        else
        {
            if (!File.Exists(ConfigPath))
            {
                throw new FileNotFoundException($"The file {ConfigPath} was not found.");
            }
            await diffConfigService.Import(ConfigPath);
        }
        diffConfigService.Reset();
        diffConfigModel.OldFileConfig.FilePath = OldPath;
        diffConfigModel.NewFileConfig.FilePath = NewPath;
        try
        {
            if (excelDiffService.SaveDiff())
            {
                console.Output.WriteLine("Diff file successfully saved.");
                return;
            }
            console.Output.WriteLine($"Saving diff file failed!");
        }
        catch (Exception e)
        {
            console.Error.WriteLine($"Error saving diff: {e.Message}");
            return;
        }

    }

}
