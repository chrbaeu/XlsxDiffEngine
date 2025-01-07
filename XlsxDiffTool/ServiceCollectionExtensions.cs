using CommunityToolkit.Mvvm.Messaging;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Serilog;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using XlsxDiffTool.Common;
using XlsxDiffTool.Models;
using XlsxDiffTool.Services;
using XlsxDiffTool.ViewModels;
using XlsxDiffTool.Views;

namespace XlsxDiffTool;

internal static class ServiceCollectionExtensions
{
    public static void AddAllServices(this IServiceCollection services)
    {
        services.AddSerilog((services, loggerConfiguration) =>
            loggerConfiguration.ReadFrom.Configuration(services.GetRequiredService<IConfiguration>()));

        services.AddSingleton<AppInfo>(new AppInfo(
            "Excel Diff UI",
            AppContext.BaseDirectory,
            Assembly.GetExecutingAssembly().GetName().Version ?? new(),
            Program.AppStartupTimestamp));

        services.AddScoped<IMessenger, WeakReferenceMessenger>();

        services.AddLocalization();

        services.AddSingleton<ViewFactory>();
        services.AddScoped<IDialogService, DialogServiceWpf>();
        services.AddScoped<WindowStateModel>();

        services.AddApplicationServices();

        services.AddViewsAndViewModels();
    }

    private static void AddApplicationServices(this IServiceCollection services)
    {
        services.AddScoped<WindowStateSettingsService>();
        services.AddScoped<DiffConfigService>();
        services.AddScoped<DiffConfigModel>();
        services.AddScoped<ColumnInfoService>();
        services.AddScoped<ExcelDiffService>();
        services.AddScoped<PluginService>();
    }

    private static void AddViewsAndViewModels(this IServiceCollection services)
    {
        services.AddViewWithViewModel<MainWindow, MainWindowViewModel>();
        services.AddViewWithViewModel<MainView, MainViewModel>();
        services.AddViewWithViewModel<InputSelectorView, OldInputSelectorViewModel>();
        services.AddViewWithViewModel<InputSelectorView, NewInputSelectorViewModel>();
        services.AddViewWithViewModel<OutputSelectorView, OutputSelectorViewModel>();
        services.AddViewWithViewModel<ColumnSelectorView, ColumnSelectorViewModel>();
        services.AddViewWithViewModel<OptionsView, OptionsViewModel>();
    }

    private static IServiceCollection AddViewWithViewModel<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicConstructors)] TView, [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicConstructors)] TViewModel>(this IServiceCollection services)
        where TView : class, IView
        where TViewModel : class, IViewModel
    {
        ViewFactory.RegisterViewModel<TViewModel, TView>();
        services.TryAddTransient<TView>();
        _ = services.AddTransient<TViewModel>();
        return services;
    }

}
