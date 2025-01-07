# XlsxDiffEngine

XlsxDiffEngine is a simple C# library for comparing Excel (.xlsx) files. It provides powerful configuration options to customize data comparisons, allowing you to generate annotated output files that highlight all changes, additions, and removals. (For reading and writing Excel files, the [EPPlus](https://github.com/EPPlusSoftware/EPPlus) library is used.)

![Logo](https://raw.githubusercontent.com/chrbaeu/XlsxDiffEngine/refs/heads/main/XlsxDiffEngine/Icon.png)

## Features

- **Flexible comparison options**: Configure key columns, secondary key columns, ignored columns, group columns, text-only comparison columns, and more.
- **Customizable data handling**: Define data ranges, set merging options, add skip rules, apply modification rules, or manage individual requirements with callbacks.
- **Visual change indicators**: Highlight changes, additions, and deletions in the output using different cell styling and comment options.
- **Configurable output**: Customize headers, auto-fit columns, freeze panes, and apply auto-filters for enhanced readability.

## Installation

Add XlsxDiffEngine to your project via NuGet:

[Chriffizient.XlsxDiffEngine on NuGet](https://www.nuget.org/packages/Chriffizient.XlsxDiffEngine)

```bash
dotnet add package Chriffizient.XlsxDiffEngine
```

## Getting Started

### Basic Usage

Use the `ExcelDiffBuilder` to set up Excel files and key columns.
Adjust other comparison options like value changed markers, ignored columns, modification rules, and styling preferences.
Use the `Build` method to save an annotated comparison Excel output file.

### Example

```csharp  
using XlsxDiffEngine;  
using OfficeOpenXml;  
  
new ExcelDiffBuilder()
    .AddFiles(x => x
        .SetOldFile(oldFileStream, "OldFile.xlsx")
        .SetNewFile(newFileStream, "NewFile.xlsx")
        )
    .SetKeyColumns("ID") // Optional
    .Build("ComparisonOutput.xlsx");
```

For more examples, take a look at the tests. Or try out the functions via the XlsxDiffTool WPF application.

## Dependencies

- [EPPlus](https://github.com/EPPlusSoftware/EPPlus) - for Excel file handling in .NET (Depending on the usage, a paid license can be required)
