namespace XlsxDiffEngineTests;

internal static class ExcelTestData
{
    public static object?[][] StandardOld() => Clone([
        ["Title", "Value"],
        ["A", 1],
        ["B", 2],
        ["C", 3],
    ]);

    public static object?[][] StandardNew() => Clone([
        ["Title", "Value"],
        ["A", 1],
        ["B", 4],
        ["C", 3],
    ]);

    public static object?[][] GroupedOld() => Clone([
        ["Title", "Group", "Value"],
        ["A", "1", 1],
        ["B", "1", 2],
        ["C", "2", 3],
        ["D", "2", 4],
    ]);

    public static object?[][] GroupedNew() => Clone([
        ["Title", "Group", "Value"],
        ["A", "1", 1],
        ["E", "1", 5],
        ["C", "2", 3],
        ["F", "2", 6],
    ]);

    public static object?[][] NumericMarkerOld() => Clone([
        ["Title", "Value"],
        ["A", 100.0],
        ["B", 100.0],
        ["C", 100.0],
        ["D", 100.0],
    ]);

    public static object?[][] NumericMarkerNew() => Clone([
        ["Title", "Value"],
        ["A", 100.5],
        ["B", 111.0],
        ["C", 130.0],
        ["D", 100.0],
    ]);

    public static object?[][] NumericRuleBase() => Clone([
        ["Title", "Value"],
        ["A", 100.00],
        ["B", 100.00],
        ["C", 100.00],
        ["D", 100.00],
    ]);

    public static object?[][] TypedValueTemplate() => Clone([
        ["Title", "Value"],
        ["A", null],
        ["B", null],
    ]);

    public static object?[][] SecondaryKeyOld() => Clone([
        ["ID", "SecondaryID", "Value"],
        ["1", "A", 100],
        ["2", "B", 200],
    ]);

    public static object?[][] SecondaryKeyNew() => Clone([
        ["ID", "SecondaryID", "Value"],
        ["3", "A", 100],
        ["4", "B", 250],
    ]);

    public static object?[][] SecondWorksheet() => Clone([
        ["Title", "Value"],
        ["D", 4],
        ["E", 5],
        ["F", 6],
    ]);

    private static object?[][] Clone(object?[][] data)
    {
        object?[][] copy = new object?[data.Length][];

        for (int row = 0; row < data.Length; row++)
        {
            copy[row] = new object?[data[row].Length];
            Array.Copy(data[row], copy[row], data[row].Length);
        }

        return copy;
    }
}
