using CommunityToolkit.Mvvm.ComponentModel;
using System.IO;

namespace XlsxDiffTool.Models;

public sealed class OldFileConfigModel() : FileConfigModel()
{
}

public sealed class NewFileConfigModel() : FileConfigModel()
{
}

public sealed partial class OutputFileConfigModel : FileConfigModel
{
    [ObservableProperty]
    public partial bool AddDateTime { get; set; } = true;

    public OutputFileConfigModel() : base()
    {
        IsFolderConfig = true;
        FilePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
    }
}

public abstract partial class FileConfigModel : ObservableObject
{
    [ObservableProperty]
    public partial string FilePath { get; set; } = "";

    [ObservableProperty]
    public partial int StartRow { get; set; } = 1;

    [ObservableProperty]
    public partial int StartColumn { get; set; } = 1;

    [ObservableProperty]
    public partial bool IsFolderConfig { get; set; } = false;

    [ObservableProperty]
    public partial string FileNameSelectorRegex { get; set; } = "";

    public bool IsValidPath() => !string.IsNullOrWhiteSpace(FilePath);

    public bool IsExisitingFile() => IsValidPath() && File.Exists(FilePath);

    //partial void OnFilePathChanged(string value) => messenger.Send<FileConfigChangedEvent>(new(this));
    //partial void OnStartRowChanged(int value) => messenger.Send<FileConfigChangedEvent>(new(this));
    //partial void OnStartColumnChanged(int value) => messenger.Send<FileConfigChangedEvent>(new(this));
}
