namespace ExcelDiffUI.Common;

public sealed record class AppInfo(
    string AppName,
    string ExePath,
    Version Version,
    long StartupTimestamp
    );
