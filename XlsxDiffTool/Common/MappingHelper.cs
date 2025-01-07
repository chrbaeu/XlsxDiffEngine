using System.Collections.Concurrent;
using System.Linq.Expressions;
using System.Reflection;

namespace XlsxDiffTool.Common;

public static class MappingHelper
{
    private static readonly ConcurrentDictionary<(Type SrcType, Type DstType), object> mapperDict = [];

    public static TDst Map<TSrc, TDst>(TSrc src, TDst? dst = default)
    {
        var mapper = mapperDict.GetOrAdd((typeof(TSrc), typeof(TDst)), _ => CreateMapping<TSrc, TDst>());
        return ((Func<TSrc, TDst?, TDst>)mapper)(src, dst);
    }

    private static Func<TSrc, TDst?, TDst> CreateMapping<TSrc, TDst>()
    {
        ParameterExpression srcParam = Expression.Parameter(typeof(TSrc), "Src");
        ParameterExpression dstParam = Expression.Parameter(typeof(TDst?), "Dst");
        List<ParameterExpression> parameters = [srcParam, dstParam];

        BinaryExpression dstParamNullCheck = Expression.ReferenceEqual(dstParam, Expression.Constant(null));
        ParameterExpression wasDstParamNullVar = Expression.Variable(typeof(bool), "wasDstParamNullVar");
        Expression assignWasDstParamNullVar = Expression.Assign(wasDstParamNullVar, dstParamNullCheck);
        List<ParameterExpression> localVariables = [wasDstParamNullVar];

        ConstructorInfo constructor = typeof(TDst).GetConstructor([]) ?? throw new NotImplementedException($"No default constructor for type {typeof(TDst).Name} found!");
        NewExpression newObject = Expression.New(constructor);
        BinaryExpression assignObject = Expression.Assign(dstParam, newObject);
        Expression asignNewObjectIfNeededExpression = Expression.IfThen(wasDstParamNullVar, assignObject);

        List<Expression> expressions = [assignWasDstParamNullVar, asignNewObjectIfNeededExpression];

        var dstPropertyDict = typeof(TDst).GetProperties().Where(x => x.CanWrite).ToDictionary(x => x.Name, StringComparer.OrdinalIgnoreCase);
        foreach (PropertyInfo? srcProperty in typeof(TSrc).GetProperties().Where(x => x.CanRead))
        {
            if (dstPropertyDict.TryGetValue(srcProperty.Name, out PropertyInfo? dstProperty))
            {
                if (srcProperty.GetMethod is null) { throw new NotImplementedException($"No getter for property {srcProperty.Name} in type {typeof(TSrc).Name} found!"); }
                if (dstProperty.SetMethod is null) { throw new NotImplementedException($"No setter for property {srcProperty.Name} in type {typeof(TDst).Name} found!"); }
                if (dstParam.Type.IsValueType) { throw new NotImplementedException("Handling of InitOnly properties in value types is not supported."); }

                MethodCallExpression getValue = Expression.Call(srcParam, srcProperty.GetMethod);
                Expression setValue = Expression.Call(dstParam, dstProperty.SetMethod, getValue);
                if (IsInitOnly(dstProperty)) { setValue = Expression.IfThen(wasDstParamNullVar, setValue); }
                expressions.Add(setValue);
            }
        }
        expressions.Add(dstParam);

        BlockExpression body = Expression.Block(localVariables, expressions);
        var expression = Expression.Lambda<Func<TSrc, TDst?, TDst>>(body, parameters);
        return expression.Compile();
    }

    private static bool IsInitOnly(PropertyInfo propertyInfo)
    {
        return propertyInfo.SetMethod?.ReturnParameter?.GetRequiredCustomModifiers().Contains(typeof(System.Runtime.CompilerServices.IsExternalInit)) ?? false;
    }
}
