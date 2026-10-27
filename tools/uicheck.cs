// uicheck.cs —— 启动安装向导，只截"向导窗口"那块区域，逐页确认中文排版
//   编译：uiprobe-build.bat
//   用法：_probe\uicheck.exe <安装包路径> <输出前缀> [页数]
//   行为：启动 -> 找到标题含 "OfficeImage" 的顶层窗口 -> 截欢迎页
//         -> Alt+N 翻页 -> 截图 ... -> 关闭（不真的安装）
//
//   注意：安装包要先自提权（Inno 的 loader 用 ShellExecute "runas"），
//         所以必须 UseShellExecute=true，否则窗口根本不会出现。
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

internal static class UiCheck
{
    private delegate bool EnumProc(IntPtr h, IntPtr l);

    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowTextW(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] private static extern bool BringWindowToTop(IntPtr h);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
    [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT { public int L, T, R, B; }

    private static string Title(IntPtr h)
    {
        StringBuilder sb = new StringBuilder(512);
        GetWindowTextW(h, sb, sb.Capacity);
        return sb.ToString();
    }

    private static IntPtr FindWindow(string needle, bool needVisible)
    {
        IntPtr found = IntPtr.Zero;
        EnumWindows(delegate(IntPtr h, IntPtr l)
        {
            if (needVisible && !IsWindowVisible(h)) return true;
            string t = Title(h);
            if (t.Length > 0 && t.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                found = h;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    private static void ListWindows()
    {
        Console.WriteLine("  可见窗口:");
        EnumWindows(delegate(IntPtr h, IntPtr l)
        {
            if (IsWindowVisible(h))
            {
                string t = Title(h);
                if (t.Length > 0) Console.WriteLine("    '" + t + "'");
            }
            return true;
        }, IntPtr.Zero);
    }

    private static void ShotWindow(IntPtr h, string path)
    {
        SetWindowPos(h, new IntPtr(-1), 0, 0, 0, 0, SWP_NOSIZE_MOVE_SHOW);
        Thread.Sleep(250);
        RECT r;
        if (!GetWindowRect(h, out r)) { Console.WriteLine("  GetWindowRect 失败"); return; }
        int w = r.R - r.L, hh = r.B - r.T;
        if (w <= 0 || hh <= 0) { Console.WriteLine("  窗口尺寸异常"); return; }
        using (Bitmap bmp = new Bitmap(w, hh))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.CopyFromScreen(r.L, r.T, 0, 0, new Size(w, hh));
            bmp.Save(path, ImageFormat.Png);
        }
        Console.WriteLine("  shot -> " + path + "  (" + w + "x" + hh + ")");
    }

    private const uint SWP_NOSIZE_MOVE_SHOW = 0x0043;   // NOSIZE|NOMOVE|SHOWWINDOW

    private static bool Focus(IntPtr h)
    {
        // 前台窗口锁：非前台进程直接 SetForegroundWindow 会被系统忽略，
        // 先把窗口顶到 TOPMOST（保证截图时不被别的窗口盖住），
        // 再 AttachThreadInput 到当前前台线程，借它的权限设置前台。
        SetWindowPos(h, new IntPtr(-1), 0, 0, 0, 0, SWP_NOSIZE_MOVE_SHOW);

        uint pid;
        uint fgThread = GetWindowThreadProcessId(GetForegroundWindow(), out pid);
        uint me = GetCurrentThreadId();
        bool attached = AttachThreadInput(me, fgThread, true);
        BringWindowToTop(h);
        SetForegroundWindow(h);
        if (attached) AttachThreadInput(me, fgThread, false);

        for (int i = 0; i < 10; i++)
        {
            Thread.Sleep(120);
            if (GetForegroundWindow() == h) return true;
        }
        return GetForegroundWindow() == h;
    }

    private static void Main(string[] args)
    {
        if (args.Length < 2) { Console.WriteLine("usage: uicheck.exe <setup.exe> <prefix> [pages]"); return; }
        string exe = args[0];
        string prefix = args[1];
        int pages = args.Length > 2 ? int.Parse(args[2]) : 3;

        ProcessStartInfo psi = new ProcessStartInfo(exe) { UseShellExecute = true };
        Process p = Process.Start(psi);
        Console.WriteLine("started");

        IntPtr hwnd = IntPtr.Zero;
        for (int i = 0; i < 60; i++)
        {
            Thread.Sleep(250);
            hwnd = FindWindow("OfficeImage", true);
            if (hwnd != IntPtr.Zero) break;
        }
        if (hwnd == IntPtr.Zero) { Console.WriteLine("没找到安装向导窗口"); ListWindows(); return; }
        Console.WriteLine("hwnd=" + hwnd + "  title='" + Title(hwnd) + "'");

        if (!Focus(hwnd)) Console.WriteLine("  !! 窗口没抢到前台，翻页可能失败");
        Thread.Sleep(1200);

        ShotWindow(hwnd, prefix + "-1.png");

        for (int i = 2; i <= pages; i++)
        {
            SendKeys.SendWait("%n");     // Alt+N = 下一步
            Thread.Sleep(1400);
            Focus(hwnd);
            Thread.Sleep(300);
            ShotWindow(hwnd, prefix + "-" + i + ".png");
        }

        PostMessage(hwnd, 0x0010, IntPtr.Zero, IntPtr.Zero);   // WM_CLOSE，取消安装
        Thread.Sleep(1200);
        if (GetForegroundWindow() == hwnd) PostMessage(hwnd, 0x0010, IntPtr.Zero, IntPtr.Zero);
        Console.WriteLine("closed");
    }
}
