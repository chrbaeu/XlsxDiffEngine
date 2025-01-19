using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Localization;
using OfficeOpenXml;
using Serilog;
using System.Globalization;
using System.Windows;
using XlsxDiffTool.Common;
using XlsxDiffTool.Services;
using XlsxDiffTool.ViewModels;
using XlsxDiffTool.Views;

namespace XlsxDiffTool;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    private readonly IHost appHost;

    public App()
    {
        HostApplicationBuilder builder = Host.CreateApplicationBuilder();

        if (builder.Configuration["Language"] is string { } language and not "")
        {
            CultureInfo cultureInfo = new(language);
            CultureInfo.DefaultThreadCurrentCulture = cultureInfo;
            CultureInfo.DefaultThreadCurrentUICulture = cultureInfo;
        }

        builder.Services.AddAllServices();

        appHost = builder.Build();
    }

    protected override async void OnStartup(StartupEventArgs e)
    {
        await appHost.StartAsync();

        // Static configurations
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        ViewConverter.ViewFactory = appHost.Services.GetRequiredService<ViewFactory>();
        TranslateExtension.Localizer = appHost.Services.GetRequiredService<IStringLocalizer<Resources.Resources>>();

        // Create main window view model
        await appHost.Services.GetRequiredService<DiffConfigService>().Reset();
        MainWindowViewModel vm = appHost.Services.GetRequiredService<MainWindowViewModel>();
        await appHost.Services.GetRequiredService<WindowStateSettingsService>().TryToRestoreWindowStateAsync(vm.WindowState);

        // Create main window
        MainWindow mainWindow = (MainWindow)ViewConverter.ViewFactory.GetOrCreateView(vm);
        mainWindow.Closed += async (s, e) =>
        {
            if (s is MainWindow { DataContext: MainWindowViewModel vm })
            {
                await appHost.Services.GetRequiredService<WindowStateSettingsService>().SaveWindowStateAsync(vm.WindowState);
            }
        };

        Log.Information("Show application window");
        mainWindow.Show();
    }

    protected override async void OnExit(ExitEventArgs e)
    {
        await appHost.StopAsync();
        appHost.Dispose();
        base.OnExit(e);
    }
}
