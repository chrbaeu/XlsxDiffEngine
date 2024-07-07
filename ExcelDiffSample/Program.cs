using ExcelDiffEngine;
using OfficeOpenXml;

namespace ExcelDiffSample;

internal class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("ExcelDiff Sample");

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var folder = Path.GetFullPath(".");
        new ExcelDiffBuilder()
            .AddFiles(x => x
                .SetOldFile(Path.Combine(folder, @"2023\0_Rohbau_2023-07-14.xlsx"))
                .SetNewFile(Path.Combine(folder, @"2024\0_Rohbau_2024-07-01.xlsx"))
                .SetMergedWorksheetName("Rohbau")
                .SetDataArea(7))
            .AddFiles(x => x
                .SetOldFile(Path.Combine(folder, @"2023\1_Ausbau_2023-07-14.xlsx"))
                .SetNewFile(Path.Combine(folder, @"2024\2_Ausbau_2024-07-01.xlsx"))
                .SetMergedWorksheetName("Ausbau")
                .SetDataArea(7))
            .AddFiles(x => x
                .SetOldFile(Path.Combine(folder, @"2023\2_Gebäudetechnik_2023-07-14.xlsx"))
                .SetNewFile(Path.Combine(folder, @"2024\3_Gebäudetechnik_2024-07-01.xlsx"))
                .SetMergedWorksheetName("Gebäudetechnik")
                .SetDataArea(7))
            .AddFiles(x => x
                .SetOldFile(Path.Combine(folder, @"2023\3_Freianlagen_2023-07-14.xlsx"))
                .SetNewFile(Path.Combine(folder, @"2024\1_Freianlagen_2024-07-01.xlsx"))
                .SetMergedWorksheetName("Freianlagen")
                .SetDataArea(7))
            .AddFiles(x => x
                .SetOldFile(Path.Combine(folder, @"2023\4_Instandsetzung___Abbruch_2023-07-14.xlsx"))
                .SetNewFile(Path.Combine(folder, @"2024\4_Instandsetzung___Abbruch_2024-07-01.xlsx"))
                .SetMergedWorksheetName("Instandsetzung - Abbruch")
                .SetDataArea(7))
            .AddWorksheetNameAsColumn("LB")
            .AddMergedWorksheetNameAsColumn("Kategorie")
            //.AddDocumentNameAsColumn("Datei")
            .MergeWorkSheets()
            .MergeDocuments()
            .SetMergedDocumentName("Vergleich")
            .SetColumsToIgnore("UUID")
            .SetKeyColumns("BKI Positionsnummer")
            .AddValueChangedMarker(0, 1, DefaultCellStyles.YellowValueChangedMarker)
            .AddValueChangedMarker(0.05, 1, DefaultCellStyles.OrangeValueChangedMarker)
            .AddValueChangedMarker(0.10, 1, DefaultCellStyles.RedValueChangedMarker)
            .Build("Diff.xlsx");
    }
}
