using Microsoft.Extensions.DependencyInjection;
using System.Collections.Concurrent;

namespace ExcelDiffUI.Common;

public sealed class ViewFactory(IServiceProvider serviceProvider)
{
    private static readonly ConcurrentDictionary<Type, Type> viewModelTypeToViewTypeDict = new();

    private readonly IServiceProvider serviceProvider = serviceProvider;
    private readonly ConcurrentDictionary<Type, MappingList> instances = new();

    public static void RegisterViewModel<TViewModel, TView>()
        where TViewModel : IViewModel
        where TView : IView
    {
        viewModelTypeToViewTypeDict[typeof(TViewModel)] = typeof(TView);
    }

    public IView? GetViewOrDefault(IViewModel viewModel)
    {
        MappingList mappingList = instances.GetOrAdd(viewModel.GetType(), _ => new());
        return mappingList.GetViewOrDefault(viewModel);
    }

    public IView GetOrCreateView(IViewModel viewModel)
    {
        MappingList mappingList = instances.GetOrAdd(viewModel.GetType(), _ => new());
        return mappingList.GetViewOrAdd(viewModel, () =>
        {
            if (viewModelTypeToViewTypeDict.TryGetValue(viewModel.GetType(), out Type? viewType))
            {
                var view = (IView)serviceProvider.GetRequiredService(viewType);
                view.DataContext = viewModel;
                return view;
            }
            throw new ArgumentException($"No view found for view model of type {viewModel.GetType()}");
        });
    }

    private sealed class MappingList
    {
        private readonly List<MappingEntry> list = [];
        private readonly object lockObject = new();

        public IView? GetViewOrDefault(IViewModel viewModel)
        {
            lock (lockObject)
            {
                return GetViewInternal(viewModel);
            }
        }

        public IView GetViewOrAdd(IViewModel viewModel, Func<IView> viewFactory)
        {
            lock (lockObject)
            {
                IView view = GetViewInternal(viewModel) ?? viewFactory();
                list.Add(new MappingEntry(new(view), new(viewModel)));
                return view;
            }
        }

        public IView? GetViewInternal(IViewModel viewModel)
        {
            bool cleanup = false;
            foreach (MappingEntry entry in list)
            {
                if (entry.View.TryGetTarget(out IView? viewValue) && entry.ViewModel.TryGetTarget(out IViewModel? viewModelValue))
                {
                    if (viewModelValue == viewModel) { return viewValue; }
                }
                else
                {
                    cleanup = true;
                }
            }
            if (cleanup)
            {
                list.RemoveAll(entry => !entry.View.TryGetTarget(out _) || !entry.ViewModel.TryGetTarget(out _));
            }
            return null;
        }

        private sealed record class MappingEntry(WeakReference<IView> View, WeakReference<IViewModel> ViewModel);
    }
}
