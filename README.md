Coming soon ...

# ExcelDiffEngine

ExcelDiffEngine is a simple C# library for  comparing Excel (.xlsx) files. It provides powerful configuration options to customize data comparisons, allowing you to generate annotated output files that highlight all changes, additions, and removals.


![Logo](https://github.com/chrbaeu/ExcelDiffEngine/blob/main/ExcelDiff/Icon.png?raw=true)

## Features

- **Flexible Comparison Options**: Configure key columns, secondary key columns, ignored columns, group columns, text-only comparison columns, and more.
- **Customizable Data Handling**: Define data ranges, set merging options, add skip rules, apply modification rules, or manage individual requirements with callbacks.
- **Visual Change Indicators**: Highlight changes, additions, and deletions in the output using different cell styling and comment options.
- **Configurable Output**: Customize headers, auto-fit columns, freeze panes, and apply auto-filters for enhanced readability.

## Installation

Add ExcelDiffEngine to your project via NuGet:

```bash
dotnet add package ExcelDiffEngine
```

## Getting Started

### Basic Usage

Use the `ExcelDiffBuilder` to set up Excel files and key columns.
Adjust other comparison options like value changed markers, ignored columns, modification rules, and styling preferences.
Use the `Build` method to save an annotated comparison Excel output file.

### Example

```csharp  
using ExcelDiffEngine;  
using OfficeOpenXml;  
  
var builder = new ExcelDiffBuilder()  
    .AddFiles(config =>  
    {  
        config  
            .SetOldFile("OldFile.xlsx")  
            .SetNewFile("NewFile.xlsx");  
    })  
    .SetKeyColumns("ID") // Set key column(s) to identify rows  
    .Build("ComparisonOutput.xlsx");  
```

## Dependencies

- [EPPlus](https://github.com/EPPlusSoftware/EPPlus) - for Excel file handling in .NET.
