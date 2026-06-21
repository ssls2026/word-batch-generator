using System;
using System.Runtime.InteropServices;
using System.IO;

class CheckIcon {
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode)]
    static extern IntPtr LoadLibraryEx(string f, IntPtr h, uint d);
    [DllImport("kernel32.dll")]
    static extern bool FreeLibrary(IntPtr h);
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode)]
    static extern IntPtr FindResource(IntPtr m, IntPtr n, IntPtr t);
    [DllImport("kernel32.dll")]
    static extern uint SizeofResource(IntPtr m, IntPtr r);

    static void Main(string[] args) {
        string exe = args.Length > 0 ? args[0] : @"dist\Word批量生成器.exe";
        Console.WriteLine($"检查: {exe}");
        Console.WriteLine($"文件大小: {new FileInfo(exe).Length / 1024 / 1024} MB");
        
        var h = LoadLibraryEx(exe, IntPtr.Zero, 2); // LOAD_LIBRARY_AS_DATAFILE
        if (h == IntPtr.Zero) {
            Console.WriteLine($"LoadLibraryEx 失败，错误: {Marshal.GetLastWin32Error()}");
            return;
        }
        
        // RT_ICON=3, RT_GROUP_ICON=14
        var r14 = FindResource(h, (IntPtr)1, (IntPtr)14); // Group icon ID=1
        var r3  = FindResource(h, (IntPtr)1, (IntPtr)3);  // Icon ID=1
        Console.WriteLine(r14 != IntPtr.Zero 
            ? $"✅ RT_GROUP_ICON 存在, size={SizeofResource(h,r14)} bytes" 
            : "❌ RT_GROUP_ICON 不存在 (图标未嵌入PE!)");
        Console.WriteLine(r3 != IntPtr.Zero 
            ? $"✅ RT_ICON 存在, size={SizeofResource(h,r3)} bytes" 
            : "❌ RT_ICON 不存在");
        
        FreeLibrary(h);
    }
}
