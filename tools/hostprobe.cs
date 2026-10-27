// 验证 ImageZoomAddin 的宿主白名单（Connect.IsSupportedHost）。
//
// 不依赖 Office、不依赖 COM 激活：直接反射调用那个私有静态方法，
// 并靠"改自己的进程名"来模拟不同宿主。
//
// 用法：
//   hostprobe.exe                        -> 当前进程名（hostprobe），期望 supported=False
//   copy hostprobe.exe wps.exe  & wps.exe      -> 期望 False（WPS 文案，应被拒绝）
//   copy hostprobe.exe WINWORD.EXE & WINWORD.EXE -> 期望 True, host=Word
//   copy hostprobe.exe EXCEL.EXE  & EXCEL.EXE   -> 期望 True, host=Excel

using System;
using System.Reflection;

internal static class HostProbe
{
    private static int Main(string[] args)
    {
        AppDomain.CurrentDomain.AssemblyResolve += delegate(object s, ResolveEventArgs e)
        {
            string simple = new AssemblyName(e.Name).Name;
            if (simple.Equals("Extensibility", StringComparison.OrdinalIgnoreCase))
            {
                string pia = @"C:\Windows\assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\extensibility.dll";
                try { return Assembly.LoadFrom(pia); } catch { return null; }
            }
            return null;
        };

        string dll = args.Length > 0
            ? args[0]
            : @"C:\Users\BX\WorkBuddy\2026-10-26-11-31-15\imagezoom-addin\ImageZoomAddin.dll";

        string proc = System.Diagnostics.Process.GetCurrentProcess().ProcessName;
        Console.WriteLine("process  : " + proc);

        Assembly asm;
        try { asm = Assembly.LoadFrom(dll); }
        catch (Exception ex) { Console.WriteLine("LOAD FAILED: " + ex.Message); return 2; }

        Type t = asm.GetType("ImageZoomAddin.Connect");
        if (t == null) { Console.WriteLine("TYPE NOT FOUND"); return 2; }

        MethodInfo m = t.GetMethod("IsSupportedHost",
            BindingFlags.NonPublic | BindingFlags.Static);
        if (m == null) { Console.WriteLine("METHOD NOT FOUND"); return 2; }

        object[] parms = new object[] { null };
        bool ok;
        try { ok = (bool)m.Invoke(null, parms); }
        catch (Exception ex) { Console.WriteLine("INVOKE FAILED: " + ex.InnerException); return 2; }

        Console.WriteLine("supported: " + ok);
        Console.WriteLine("host     : " + (parms[0] ?? "(none)"));

        bool expectTrue =
            proc.Equals("WINWORD", StringComparison.OrdinalIgnoreCase) ||
            proc.Equals("EXCEL", StringComparison.OrdinalIgnoreCase);

        Console.WriteLine(ok == expectTrue ? "RESULT: PASS" : "RESULT: FAIL");
        return ok == expectTrue ? 0 : 1;
    }
}
