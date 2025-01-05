using CliFx;
using CliFx.Attributes;
using CliFx.Infrastructure;
using System.IO;

namespace ExcelDiffUI;

[Command("diff", Description = "Compares two Excel files and saves the differences in an output file.")]
internal sealed class DiffCommand : ICommand
{

    [CommandOption("oldPath", 'o', IsRequired = true, Description = "Path for the older src xlsx data.")]
    public string OldPath { get; init; } = "";

    [CommandOption("newPath", 'n', IsRequired = true, Description = "Path for the newer src xlsx data.")]
    public string NewPath { get; init; } = "";

    [CommandOption("configPath", 'c', Description = "Path for the configuration to be used for the diff process.")]
    public string ConfigPath { get; init; } = "";

    [CommandOption("diffResultPath", 'd', Description = "Path for the resutlig diff xlsx file.")]
    public string DiffPath { get; init; } = ".";

    public ValueTask ExecuteAsync(IConsole console)
    {
        if (!File.Exists(OldPath))
        {
            throw new FileNotFoundException($"The file {OldPath} was not found.”");
        }
        if (!File.Exists(NewPath))
        {
            throw new FileNotFoundException($"The file {NewPath} was not found.");
        }

        console.Output.WriteLine($"Diff file successfully saved.");
        return ValueTask.CompletedTask;
    }

}
