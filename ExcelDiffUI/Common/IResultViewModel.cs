namespace ExcelDiffUI.Common;

public interface IResultViewModel<out T> : IViewModel
{
    public T Result { get; }
}
