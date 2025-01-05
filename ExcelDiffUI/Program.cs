using CliFx;
using ExcelDiffUI.Common;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace ExcelDiffUI;

internal static class Program
{
    public static long AppStartupTimestamp { get; private set; }

    [STAThread]
    public static int Main(string[] args)
    {
        AppStartupTimestamp = Stopwatch.GetTimestamp();
        return RunApp(args);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static int RunApp(string[] args)
    {
        if (args.Length > 0)
        {
            ConsoleHelper.InitConsole();
            return new CliApplicationBuilder()
                .AddCommandsFromThisAssembly()
                .Build()
                .RunAsync()
                .AsTask()
                .GetAwaiter()
                .GetResult();
        }
        else
        {
            var app = new App();
            app.InitializeComponent();
            app.Run();
            return 0;
        }
    }
}
