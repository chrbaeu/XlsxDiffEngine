using System.IO;
using System.Runtime.InteropServices;

namespace ExcelDiffUI.Common;

internal static partial class ConsoleHelper
{
    public static void InitConsole()
    {
        const int ATTACH_PARENT_PROCESS = -1;
        if (!AttachConsole(ATTACH_PARENT_PROCESS))
        {
            AllocConsole();
        }
        var stdout = Console.OpenStandardOutput();
        var writer = new StreamWriter(stdout) { AutoFlush = true };
        Console.SetOut(writer);
    }

    [LibraryImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool AttachConsole(int dwProcessId);

    [LibraryImport("kernel32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool AllocConsole();

}
