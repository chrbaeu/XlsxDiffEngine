namespace XlsxDiffEngine;

internal static class SortExtensions
{
    public static IOrderedEnumerable<T> OrderByList<T>(this IEnumerable<T> source, Func<T, List<object?>> keySelector)
    {
        return OrderByRecursive(source, keySelector, 0);
    }

    private static IOrderedEnumerable<T> OrderByRecursive<T>(IEnumerable<T> source, Func<T, List<object?>> keySelector, int index)
    {
        if (index == 0)
        {
            return source.OrderBy(x => keySelector(x)[index]);
        }
        return OrderByRecursive(source, keySelector, index - 1).ThenBy(x => keySelector(x)[index]);
    }
}
