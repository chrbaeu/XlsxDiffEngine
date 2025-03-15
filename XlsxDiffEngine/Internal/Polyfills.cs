using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

namespace XlsxDiffEngine;

internal static class ArgumentNullThrowHelper
{
    /// <summary>Throws an <see cref="ArgumentNullException"/> if <paramref name="argument"/> is null.</summary>
    /// <param name="argument">The reference type argument to validate as non-null.</param>
    /// <param name="paramName">The name of the parameter with which <paramref name="argument"/> corresponds.</param>
    public static void ThrowIfNull([NotNull] object? argument, [CallerArgumentExpression(nameof(argument))] string? paramName = null)
    {
        if (argument is null)
        {
            throw new ArgumentNullException(paramName);
        }
    }
}

#if NETSTANDARD2_0

internal sealed class UnreachableException : Exception
{
    /// <summary>
    /// Initializes a new instance of the <see cref="UnreachableException"/> class with the default error message.
    /// </summary>
    public UnreachableException()
        : base("The program executed an instruction that was thought to be unreachable.")
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="UnreachableException"/>
    /// class with a specified error message.
    /// </summary>
    /// <param name="message">The error message that explains the reason for the exception.</param>
    public UnreachableException(string? message)
        : base(message)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="UnreachableException"/>
    /// class with a specified error message and a reference to the inner exception that is the cause of
    /// this exception.
    /// </summary>
    /// <param name="message">The error message that explains the reason for the exception.</param>
    /// <param name="innerException">The exception that is the cause of the current exception.</param>
    public UnreachableException(string? message, Exception? innerException)
        : base(message, innerException)
    {
    }
}

internal static class IEnumerableExtensions
{
    public static HashSet<TSource> ToHashSet<TSource>(this IEnumerable<TSource> source) => source.ToHashSet(comparer: null);

    public static HashSet<TSource> ToHashSet<TSource>(this IEnumerable<TSource> source, IEqualityComparer<TSource>? comparer)
    {
        ArgumentNullThrowHelper.ThrowIfNull(source);
        return new HashSet<TSource>(source, comparer);
    }

    public static IEnumerable<TSource> DistinctBy<TSource, TKey>(
        this IEnumerable<TSource> source,
        Func<TSource, TKey> keySelector)
    {
        HashSet<TKey> seenKeys = new HashSet<TKey>();
        foreach (var element in source)
        {
            if (seenKeys.Add(keySelector(element)))
            {
                yield return element;
            }
        }
    }
}

internal static class StringExtensions
{
    public static bool Contains(this string s, string value, StringComparison comparisonType)
    {
        return s.IndexOf(value, comparisonType) >= 0;
    }

    public static string Replace(this string s, string oldValue, string? newValue, StringComparison comparisonType)
    {
        return comparisonType switch
        {
            StringComparison.Ordinal => s.Replace(oldValue, newValue),
            _ => throw new ArgumentException("The string comparison type passed in is currently not supported.", nameof(comparisonType)),
        };
    }

}

#endif
