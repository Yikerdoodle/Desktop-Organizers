# Prints the names of the icons Explorer is actually showing on the desktop right now.
# (Used to check that the Spotlight "Learn about this image" icon is really gone.)
Add-Type -TypeDefinition @'
using System; using System.Collections.Generic; using System.Runtime.InteropServices; using System.Text;
public static class DesktopIcons {
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern IntPtr FindWindow(string c, string t);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern IntPtr FindWindowEx(IntPtr p, IntPtr a, string c, string t);
  [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, int m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("kernel32.dll")] static extern IntPtr OpenProcess(uint a, bool i, uint pid);
  [DllImport("kernel32.dll")] static extern IntPtr VirtualAllocEx(IntPtr p, IntPtr a, UIntPtr s, uint t, uint pr);
  [DllImport("kernel32.dll")] static extern bool VirtualFreeEx(IntPtr p, IntPtr a, UIntPtr s, uint t);
  [DllImport("kernel32.dll")] static extern bool ReadProcessMemory(IntPtr p, IntPtr a, byte[] b, int n, out IntPtr r);
  [DllImport("kernel32.dll")] static extern bool WriteProcessMemory(IntPtr p, IntPtr a, byte[] b, int n, out IntPtr r);
  [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
  static IntPtr FindListView() {
    IntPtr prog = FindWindow("Progman", null);
    IntPtr def = FindWindowEx(prog, IntPtr.Zero, "SHELLDLL_DefView", null);
    if (def == IntPtr.Zero) {   // when a slideshow/spotlight wallpaper is active the view can live under a WorkerW
      IntPtr w = IntPtr.Zero;
      while ((w = FindWindowEx(IntPtr.Zero, w, "WorkerW", null)) != IntPtr.Zero) {
        def = FindWindowEx(w, IntPtr.Zero, "SHELLDLL_DefView", null);
        if (def != IntPtr.Zero) break;
      }
    }
    return FindWindowEx(def, IntPtr.Zero, "SysListView32", null);
  }
  public static List<string> Names() {
    var result = new List<string>();
    IntPtr lv = FindListView(); if (lv == IntPtr.Zero) return result;
    uint pid; GetWindowThreadProcessId(lv, out pid);
    IntPtr proc = OpenProcess(0x0008 | 0x0010 | 0x0020, false, pid);   // VM_OPERATION | VM_READ | VM_WRITE
    if (proc == IntPtr.Zero) return result;
    int count = (int)SendMessage(lv, 0x1004, IntPtr.Zero, IntPtr.Zero);   // LVM_GETITEMCOUNT
    int is64 = IntPtr.Size;
    int lvitemSize = is64 == 8 ? 88 : 60;
    IntPtr mem = VirtualAllocEx(proc, IntPtr.Zero, (UIntPtr)4096, 0x3000, 0x04);
    try {
      for (int i = 0; i < count; i++) {
        byte[] item = new byte[lvitemSize];
        BitConverter.GetBytes(0x0001).CopyTo(item, 0);          // mask = LVIF_TEXT
        BitConverter.GetBytes(i).CopyTo(item, 4);                // iItem
        long textPtr = (long)mem + 512;
        if (is64 == 8) { BitConverter.GetBytes(textPtr).CopyTo(item, 24); BitConverter.GetBytes(256).CopyTo(item, 32); }
        else { BitConverter.GetBytes((int)textPtr).CopyTo(item, 20); BitConverter.GetBytes(256).CopyTo(item, 24); }
        IntPtr n; WriteProcessMemory(proc, mem, item, item.Length, out n);
        SendMessage(lv, 0x1073, (IntPtr)i, mem);                 // LVM_GETITEMTEXTW
        byte[] buf = new byte[512]; ReadProcessMemory(proc, (IntPtr)textPtr, buf, 512, out n);
        string s = Encoding.Unicode.GetString(buf); int z = s.IndexOf('\0'); if (z >= 0) s = s.Substring(0, z);
        result.Add(s);
      }
    } finally { VirtualFreeEx(proc, mem, UIntPtr.Zero, 0x8000); CloseHandle(proc); }
    return result;
  }
}
'@
[DesktopIcons]::Names()
