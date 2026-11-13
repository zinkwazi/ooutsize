<#
.SYNOPSIS
    OOutSize shows what is using the space on a drive or folder. A folder tree with size bars sits
    next to a treemap you can drill into.

.DESCRIPTION
    Runs on what ships with Windows: Windows PowerShell and the .NET Framework. Nothing to install.
    The window is written in C# further down this file and Add-Type compiles it when the script starts.

    Scanning only lists folders (names, sizes, dates). It never opens or changes a file and it
    writes nothing to the disk it scans.

    Nothing is deleted unless you pick Move to Recycle Bin or Delete permanently from the right
    click menu and then answer Yes. Windows system files cannot be deleted from here at all.
    Start it with -ReadOnly to take the delete options away.

    It runs without admin rights, but then any folder the account cannot read is left out of the
    totals. An orange strip across the window and a console warning say so when that is the case.

    How to read the numbers:
      - Sizes are what Explorer calls Size. Sparse, compressed and cloud placeholder files are the
        exception and count as the space they really take on disk.
      - Junctions and symbolic links are not followed.
      - Hard linked files count once per name, the same as Explorer, so C:\Windows\WinSxS reads high.
      - When you scan a whole drive, the summary says how much used space the scan could not see.
        That is usually shadow copies, folders it was not allowed into and file system overhead.

    Mouse and keys:
      - Click a folder in the tree to show it in the map. Arrow keys work in the tree.
      - Double click a block in the map to zoom in. Backspace or the Up button goes back out.
      - Right click anything for Open in Explorer, Copy path, Copy listing, Scan again,
        Move to Recycle Bin and Delete permanently. Del and Shift+Del do the last two from the tree.
      - Point at a file type under the map to fade out the rest. Click it to keep it that way.
      - Hover over a block for its name, size and date. On a Windows system file the tip also says why
        it should not be deleted by hand.
      - A date in an orange tag means nothing in that folder has changed for over a year.
      - F5 scans again. Esc stops a scan and shows what it has so far.

.PARAMETER Path
    Drive or folder to scan straight away, for example C:\ or D:\Shares or \\server\share.
    Leave it off to pick from a list of drives.

.PARAMETER BackupPrivilege
    Turn on SeBackupPrivilege for this process so folders an administrator is normally locked out of
    can still be listed, such as System Volume Information and other people's profiles.
    Needs an elevated window. Off unless you ask for it.

.PARAMETER MaxFilesPerFolder
    How many files in one folder get a row of their own, biggest first. The rest of that folder
    shows as a single "smaller files" row. 0 means no limit. Default 5000.

.PARAMETER Threads
    How many threads list folders at the same time. 0, the default, picks twice the number of
    processors, at least 4 and at most 32. Use a low number such as 2 to go easy on a busy file server.

.PARAMETER ClassicScan
    List folders the way the first version did. It is slower. Try it if the totals look wrong on
    an unusual file system or network share.

.PARAMETER ReadOnly
    Leave out the delete options, for a window that can only look.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\OOutSize.ps1 C:\

.EXAMPLE
    .\OOutSize.ps1 -Path D:\Shares -BackupPrivilege

.EXAMPLE
    .\OOutSize.ps1 -Path \\server\share -Threads 4 -ReadOnly
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Path = '',

    [switch]$BackupPrivilege,

    [ValidateRange(0, 1000000)]
    [int]$MaxFilesPerFolder = 5000,

    [ValidateRange(0, 64)]
    [int]$Threads = 0,

    [switch]$ClassicScan,

    [switch]$ReadOnly
)

$ErrorActionPreference = 'Stop'

if (-not $env:SystemRoot) { throw 'OOutSize only runs on Windows.' }

# a relative path means relative to where the shell is, not where the process thinks it is
if ($Path) {
    try { $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path) }
    catch { }    # bad drive letter or the like - the window will say so
}

# the c# below is written for the compiler and winforms that come with windows powershell
# so when this gets started from pwsh 7 hand the whole thing over
if ($PSVersionTable.PSEdition -eq 'Core') {
    $winPs = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not $PSCommandPath -or -not (Test-Path -LiteralPath $winPs)) {
        throw 'OOutSize needs Windows PowerShell. Save it as a .ps1 file and run it with powershell.exe.'
    }
    $handOff = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath, '-MaxFilesPerFolder', $MaxFilesPerFolder)
    if ($Path) { $handOff += @('-Path', $Path) }
    if ($BackupPrivilege) { $handOff += '-BackupPrivilege' }
    if ($Threads) { $handOff += @('-Threads', $Threads) }
    if ($ClassicScan) { $handOff += '-ClassicScan' }
    if ($ReadOnly) { $handOff += '-ReadOnly' }
    & $winPs @handOff
    return
}

# applocker / wdac put powershell in constrained mode and that blocks add-type - say so up front
$mode = $ExecutionContext.SessionState.LanguageMode
if ($mode -ne 'FullLanguage') {
    throw "PowerShell is running in $mode mode here, which blocks the Add-Type call OOutSize is built on."
}

# no admin rights means the scan only sees what this account can read - the window says so too
try {
    $me = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    if (-not $me.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Warning 'Not running as administrator. Folders this account cannot read are left out, so the totals may be short.'
    }
}
catch { }

$source = @'
// ooutsize - folder sizes as a tree with bars next to a treemap you can drill into
// plain .net framework 4.5 and c# 5 only so the compiler that ships with windows can build it
// no linq and nothing from system.core on purpose - one less reference to go wrong
#pragma warning disable 0649, 0169, 0414

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace OOutSize
{
    internal delegate void NodeEvent(Node n);
    internal delegate void NodeAtEvent(Node n, Point screen);
    internal delegate void TextEvent(string s);
    internal delegate void SlotEvent(int slot);

    // one thing on disk - a folder, a file, or a pile of small files rolled into one row
    internal sealed class Node
    {
        public const byte KDir = 0, KFile = 1, KPile = 2;
        public string Name;
        public Node Parent;
        public Node[] Kids;      // null on files and piles
        public long Size;
        public long Stamp;       // last write as a filetime - newest one inside for folders
        public int Files;
        public int Dirs;
        public short Type;       // slot in the extension table, -1 when it has none
        public byte Kind;
        public bool Denied;
        public bool Open;        // expanded in the tree
        public byte Guard;       // not 0 means windows owns it and it should not be deleted by hand, see Guard
    }

    internal static class Tree
    {
        public static readonly Node[] None = new Node[0];
        public static readonly char Sep = Path.DirectorySeparatorChar;

        public static string PathOf(Node n)
        {
            if (n == null) return "";
            if (n.Kind == Node.KPile) n = n.Parent;
            List<string> parts = new List<string>();
            for (Node p = n; p != null; p = p.Parent) parts.Add(p.Name);
            StringBuilder sb = new StringBuilder(260);
            for (int i = parts.Count - 1; i >= 0; i--)
            {
                if (sb.Length > 0)
                {
                    char last = sb[sb.Length - 1];
                    if (last != '\\' && last != '/') sb.Append(Sep);
                }
                sb.Append(parts[i]);
            }
            return sb.ToString();
        }

        public static bool HasKids(Node n)
        {
            return n != null && n.Kind == Node.KDir && n.Kids != null && n.Kids.Length > 0;
        }

        public static bool Under(Node n, Node top)
        {
            for (; n != null; n = n.Parent) if (n == top) return true;
            return false;
        }

        public static int BigFirst(Node a, Node b)
        {
            if (a.Size != b.Size) return a.Size > b.Size ? -1 : 1;
            return string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);
        }

        // a folder is the sum of its rows. sorts them big first while it is there
        public static void AddUp(Node d, Comparison<Node> order)
        {
            if (d.Kids == null) d.Kids = None;
            Node[] kids = d.Kids;
            long size = 0, newest = 0, files = 0, dirs = 0;
            for (int k = 0; k < kids.Length; k++)
            {
                Node c = kids[k];
                size += c.Size;
                files += c.Files;
                if (c.Kind == Node.KDir) dirs += 1 + c.Dirs;
                if (c.Stamp > newest) newest = c.Stamp;
            }
            d.Size = size;
            d.Stamp = newest;
            d.Files = files > int.MaxValue ? int.MaxValue : (int)files;
            d.Dirs = dirs > int.MaxValue ? int.MaxValue : (int)dirs;
            if (kids.Length > 1) Array.Sort<Node>(kids, order);
        }

        // after a row is taken out or swapped, redo that folder and every folder above it
        public static void Recount(Node from)
        {
            Comparison<Node> order = new Comparison<Node>(BigFirst);
            for (Node d = from; d != null; d = d.Parent) AddUp(d, order);
        }
    }

    // the stuff on a windows volume that looks like a juicy thing to delete and is not. matched on the path
    // from the top of the volume, and everything underneath a match carries the same flag
    internal static class Guard
    {
        public static readonly string[] Advice = {
            "",
            "Part of Windows. Do not delete anything in here by hand.",
            "Windows component store. Do not delete by hand. To shrink it run DISM /Online /Cleanup-Image /StartComponentCleanup.",
            "Installer cache. Windows needs these files to repair, patch and uninstall programs. Do not delete by hand.",
            "Windows Update cache. Do not delete by hand. Clear it with Disk Cleanup, Windows Update Cleanup.",
            "Installed programs. Remove a program with Apps and Features, not by deleting its files.",
            "Windows and Microsoft application data. Do not delete by hand.",
            "Installer cache used to repair and uninstall programs. Do not delete by hand.",
            "Restore points and shadow copies. Do not delete by hand. Manage them with vssadmin or System Protection.",
            "Recycle Bin storage. Empty the Recycle Bin instead of deleting in here.",
            "Boot and recovery files. Do not delete.",
            "Windows paging file. Do not delete. Its size is set under System Properties, Advanced, Performance, Virtual memory.",
            "Hibernation file. Do not delete. Turning hibernation off with powercfg /h off removes it.",
            "Left over from a Windows upgrade. Remove it with Disk Cleanup, Clean up system files, rather than by hand."
        };

        static readonly Dictionary<string, byte> rules = Build();

        // windows itself, the page and hibernation files, shadow copies and boot files - delete refuses these.
        // program files, programdata, the recycle bin and upgrade leftovers keep the warning but can be deleted
        // after a sterner question. none of it is drawn any different, the mouseover is what te;ls you
        public static bool HandsOff(byte id)
        {
            return id != 0 && id != 5 && id != 6 && id != 7 && id != 9 && id != 13;
        }

        static Dictionary<string, byte> Build()
        {
            Dictionary<string, byte> r = new Dictionary<string, byte>(StringComparer.OrdinalIgnoreCase);
            r["Windows"] = 1;
            r["Windows\\WinSxS"] = 2;
            r["Windows\\Installer"] = 3;
            r["Windows\\SoftwareDistribution"] = 4;
            // temp files, logs and crash dumps. inside windows but fair game, so they drop out of the rule above them
            r["Windows\\Temp"] = 0;
            r["Windows\\Logs"] = 0;
            r["Windows\\Minidump"] = 0;
            r["Windows\\LiveKernelReports"] = 0;
            r["Windows\\MEMORY.DMP"] = 0;
            r["Windows.old"] = 13;
            r["$Windows.~BT"] = 13;
            r["$Windows.~WS"] = 13;
            r["Program Files"] = 5;
            r["Program Files (x86)"] = 5;
            r["ProgramData\\Microsoft"] = 6;
            r["ProgramData\\Package Cache"] = 7;
            r["System Volume Information"] = 8;
            r["$Recycle.Bin"] = 9;
            r["Recovery"] = 10;
            r["Boot"] = 10;
            r["EFI"] = 10;
            r["$WinREAgent"] = 10;
            r["bootmgr"] = 10;
            r["BOOTNXT"] = 10;
            r["pagefile.sys"] = 11;
            r["swapfile.sys"] = 11;
            r["hiberfil.sys"] = 12;
            return r;
        }

        // path below the top of the volume, backslashes, no leading or trailing one. off windows the scan root stands in for the volume
        public static string Below(string path, string scanRoot)
        {
            string p = path.Replace('/', '\\');
            int cut = 0;
            if (!Native.IsWindows) cut = Math.Min(p.Length, scanRoot.Length);
            else if (p.StartsWith("\\\\", StringComparison.Ordinal))
            {
                int a = p.IndexOf('\\', 2);
                int b = a < 0 ? -1 : p.IndexOf('\\', a + 1);
                cut = b < 0 ? p.Length : b;
            }
            else if (p.Length >= 2 && p[1] == ':') cut = 2;
            return p.Substring(cut).Trim('\\');
        }

        // rules are one or two folders deep, so only the first two levels ever need a lookup
        public static bool Shallow(string below)
        {
            return below.IndexOf('\\') < 0;
        }

        public static byte Child(string parentBelow, string name, byte inherited)
        {
            byte hit;
            return rules.TryGetValue(parentBelow.Length == 0 ? name : parentBelow + "\\" + name, out hit) ? hit : inherited;
        }

        // the windows drive without the slash, C: on nearly every box. empty off windows
        static readonly string Home = HomeDrive();

        static string HomeDrive()
        {
            try { return (Path.GetPathRoot(Environment.SystemDirectory) ?? "").TrimEnd('\\', '/'); }
            catch { return ""; }
        }

        static string OnHome(string path)
        {
            if (Home.Length == 0 || path.Length <= Home.Length + 1) return null;
            if (!path.StartsWith(Home, StringComparison.OrdinalIgnoreCase) || path[Home.Length] != '\\') return null;
            return path.Substring(Home.Length + 1);
        }

        // folders on the windows drive that are never deleted whole, though things inside them can be.
        // these only matter to delete, they do not mark anything in the map
        public static string Keep(string path)
        {
            string rel = OnHome(path);
            if (rel == null) return "";
            if (string.Equals(rel, "Users", StringComparison.OrdinalIgnoreCase)) return "Every user profile on this computer lives in this folder.";
            if (string.Equals(rel, "ProgramData", StringComparison.OrdinalIgnoreCase)) return "Windows and most installed programs keep their settings and data in this folder.";
            if (string.Equals(rel, "ProgramData\\Microsoft", StringComparison.OrdinalIgnoreCase)) return "Windows keeps its own application data in this folder.";
            if (string.Equals(rel, "Program Files", StringComparison.OrdinalIgnoreCase) || string.Equals(rel, "Program Files (x86)", StringComparison.OrdinalIgnoreCase))
                return "Every installed program lives in this folder. Remove programs one at a time with Apps and Features.";
            return "";
        }

        // allowed, but worth a line in the confirm box
        public static string Care(string path)
        {
            string rel = OnHome(path);
            if (rel == null) return "";
            if (rel.StartsWith("Users\\", StringComparison.OrdinalIgnoreCase) && rel.IndexOf('\\', 6) < 0)
                return "This is a user profile. Windows keeps its own record of each profile, so the clean way to remove one is System Properties, Advanced, User Profiles.";
            return "";
        }

        // for the folder a scan starts in, which may already be inside a guarded place
        public static byte Whole(string below)
        {
            byte found = 0, hit;
            int slash = below.IndexOf('\\');
            if (rules.TryGetValue(slash < 0 ? below : below.Substring(0, slash), out hit)) found = hit;
            if (slash >= 0)
            {
                int next = below.IndexOf('\\', slash + 1);
                if (rules.TryGetValue(next < 0 ? below : below.Substring(0, next), out hit)) found = hit;
            }
            return found;
        }
    }

    internal static class Fmt
    {
        static readonly string[] Units = { "B", "KB", "MB", "GB", "TB", "PB" };

        public static string Size(long bytes)
        {
            if (bytes < 0) return "-" + Size(-bytes);
            if (bytes < 1024) return bytes.ToString(CultureInfo.CurrentCulture) + " B";
            double v = bytes;
            int u = 0;
            while (v >= 1024 && u < Units.Length - 1) { v /= 1024; u++; }
            string f = v >= 100 ? "0" : (v >= 10 ? "0.0" : "0.00");
            return v.ToString(f, CultureInfo.CurrentCulture) + " " + Units[u];
        }

        public static string Pct(double share)
        {
            if (share <= 0) return "0%";
            double p = share * 100;
            if (p < 0.1) return "<0.1%";
            if (p >= 99.95) return "100%";
            return p.ToString("0.0", CultureInfo.CurrentCulture) + "%";
        }

        public static string Count(long n)
        {
            return n.ToString("N0", CultureInfo.CurrentCulture);
        }

        public static string Date(long stamp)
        {
            if (stamp <= 0) return "";
            try { return DateTime.FromFileTimeUtc(stamp).ToLocalTime().ToString("yyyy-MM-dd", CultureInfo.InvariantCulture); }
            catch { return ""; }
        }

        public static string Secs(double s)
        {
            if (s < 10) return s.ToString("0.0", CultureInfo.CurrentCulture) + " s";
            if (s < 90) return ((int)s).ToString(CultureInfo.CurrentCulture) + " s";
            int m = (int)(s / 60);
            return m + " min " + ((int)s - m * 60) + " s";
        }

        public static string Things(long n, string one, string many)
        {
            return Count(n) + " " + (n == 1 ? one : many);
        }
    }

    // straight win32 and one ntdll call. none of it cares about the 260 char path limit
    // as long as the path wears the \\?\ prefix
    internal static class Native
    {
        public const uint DIRECTORY = 0x10, SPARSE = 0x200, REPARSE = 0x400, COMPRESSED = 0x800,
            OFFLINE = 0x1000, RECALL_OPEN = 0x40000, RECALL_DATA = 0x400000;
        public const uint THIN = SPARSE | COMPRESSED | OFFLINE | RECALL_OPEN | RECALL_DATA;
        // set in the reparse tag of anything that stands in for some other place - junctions, mounted volumes,
        // symlinks. a tag without it, like the ones onedrive and dedup put on folders, is sti;l a real folder
        public const uint SURROGATE = 0x20000000;
        public const int NO_MORE_FILES = unchecked((int)0x80000006), NO_SUCH_FILE = unchecked((int)0xC000000F);
        public static readonly IntPtr Bad = new IntPtr(-1);
        public static readonly bool IsWindows = Environment.OSVersion.Platform == PlatformID.Win32NT;

        // WIN32_FIND_DATAW, 592 bytes
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct FindData
        {
            public uint Attr;
            public uint CreateLo, CreateHi, AccessLo, AccessHi, WriteLo, WriteHi;
            public uint SizeHi, SizeLo;
            public uint Tag;         // reparse tag when Attr says reparse point
            public uint Reserved1;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string Name;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)] public string ShortName;
        }

        // IO_STATUS_BLOCK - Information is how many bytes of the buffer got filled
        [StructLayout(LayoutKind.Sequential)]
        public struct IoStatus { public IntPtr Status; public IntPtr Information; }

        [StructLayout(LayoutKind.Sequential)]
        public struct Luid { public uint Low; public int High; }

        [StructLayout(LayoutKind.Sequential)]
        public struct TokenPrivileges { public uint Count; public Luid Id; public uint Attributes; }

        [DllImport("kernel32.dll", EntryPoint = "FindFirstFileExW", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr FindFirst(string pattern, int infoLevel, out FindData data, int searchOp, IntPtr filter, int flags);

        [DllImport("kernel32.dll", EntryPoint = "FindNextFileW", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool FindNext(IntPtr handle, out FindData data);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool FindClose(IntPtr handle);

        [DllImport("kernel32.dll", EntryPoint = "GetCompressedFileSizeW", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern uint CompressedSize(string path, out uint high);

        [DllImport("kernel32.dll", EntryPoint = "CreateFileW", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern IntPtr CreateFile(string path, uint access, uint share, IntPtr security, uint disposition, uint flags, IntPtr template);

        // class 2 is FILE_FULL_DIR_INFORMATION, the same rows findfirstfile reads when asked for the basic info level
        [DllImport("ntdll.dll")]
        public static extern int NtQueryDirectoryFile(IntPtr handle, IntPtr signal, IntPtr apcRoutine, IntPtr apcContext,
            out IoStatus io, [Out] byte[] buffer, int length, int infoClass,
            [MarshalAs(UnmanagedType.U1)] bool single, IntPtr mask, [MarshalAs(UnmanagedType.U1)] bool restart);

        // list access only, shared with everyone, and backup semantics because that is how a folder gets opened at all
        public static IntPtr OpenDir(string longPath)
        {
            return CreateFile(longPath, 0x00100001, 7, IntPtr.Zero, 3, 0x02000000, IntPtr.Zero);
        }

        [DllImport("kernel32.dll", EntryPoint = "GetDiskFreeSpaceExW", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DiskFree(string dir, out ulong freeToCaller, out ulong total, out ulong free);

        [DllImport("kernel32.dll")]
        static extern IntPtr GetCurrentProcess();

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool CloseHandle(IntPtr handle);

        [DllImport("advapi32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool OpenProcessToken(IntPtr process, uint access, out IntPtr token);

        [DllImport("advapi32.dll", EntryPoint = "LookupPrivilegeValueW", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool LookupPrivilege(string system, string name, out Luid id);

        [DllImport("advapi32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AdjustTokenPrivileges(IntPtr token, [MarshalAs(UnmanagedType.Bool)] bool disableAll,
            ref TokenPrivileges state, int length, IntPtr previous, IntPtr returned);

        [DllImport("kernel32.dll", EntryPoint = "GetFileAttributesW", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern uint Attributes(string path);

        // SHFILEOPSTRUCTW. the 32 bit headers pack it to one byte and the 64 bit ones do not, so there are two
        [StructLayout(LayoutKind.Sequential)]
        struct FileOp64
        {
            public IntPtr Owner;
            public uint Func;
            public IntPtr From;
            public IntPtr To;
            public ushort Flags;
            public int Aborted;
            public IntPtr Mappings;
            public IntPtr Title;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        struct FileOp32
        {
            public IntPtr Owner;
            public uint Func;
            public IntPtr From;
            public IntPtr To;
            public ushort Flags;
            public int Aborted;
            public IntPtr Mappings;
            public IntPtr Title;
        }

        [DllImport("shell32.dll", EntryPoint = "SHFileOperationW")]
        static extern int FileOperation64(ref FileOp64 op);

        [DllImport("shell32.dll", EntryPoint = "SHFileOperationW")]
        static extern int FileOperation32(ref FileOp32 op);

        static readonly char[] Wild = { '*', '?', '"', '<', '>', '|' };

        // missing means windows said not found. anything else, access denied included, counts as sti;l there
        public static bool Exists(string path)
        {
            if (!IsWindows) return File.Exists(path) || Directory.Exists(path);
            if (Attributes(Long(path)) != 0xFFFFFFFF) return true;
            int err = Marshal.GetLastWin32Error();
            return err != 2 && err != 3;
        }

        // the shell tidies a path before it acts on it. one that would change under that tidy up - a name ending in a dot
        // or a space, a device name, wildcards, over 259 characters - could land on a different item, so those are refused
        public static bool Plain(string path)
        {
            if (path.Length > 259 || path.IndexOfAny(Wild) >= 0) return false;
            try { return string.Equals(Path.GetFullPath(path), path, StringComparison.Ordinal); }
            catch { return false; }
        }

        // hands the delete to the windows shell, the code explorer itself runs. it knows about the recycle bin, junctions,
        // files in use and when to ask for elevation, and it puts up its own progress and error windows. 0 means it ran
        public static int ShellDelete(IntPtr owner, string path, bool recycle, out bool aborted)
        {
            aborted = false;
            if (!IsWindows) return 50;
            // the shell wants a list that ends in two nulls. one goes on here and the marshaller adds the other
            IntPtr from = Marshal.StringToHGlobalUni(path + "\0");
            try
            {
                // 0x10 skip the shell's own are you sure, ours came first. 0x2000 only what was named, so a web page does
                // not take its _files folder along. 0x8000 a link is a thing to remove, never a door to walk through.
                // recycle adds 0x40 allow undo, and 0x4000 so it sti;l asks when an item cannot go to the bin
                ushort flags = (ushort)(recycle ? 0x10 | 0x2000 | 0x8000 | 0x40 | 0x4000 : 0x10 | 0x2000 | 0x8000);
                int code;
                if (IntPtr.Size == 8)
                {
                    FileOp64 op = new FileOp64();
                    op.Owner = owner;
                    op.Func = 3;        // FO_DELETE
                    op.From = from;
                    op.Flags = flags;
                    code = FileOperation64(ref op);
                    aborted = op.Aborted != 0;
                }
                else
                {
                    FileOp32 op = new FileOp32();
                    op.Owner = owner;
                    op.Func = 3;
                    op.From = from;
                    op.Flags = flags;
                    code = FileOperation32(ref op);
                    aborted = op.Aborted != 0;
                }
                return code;
            }
            finally { Marshal.FreeHGlobal(from); }
        }

        public static string Long(string path)
        {
            if (path.StartsWith(@"\\?\", StringComparison.Ordinal)) return path;
            if (path.StartsWith(@"\\", StringComparison.Ordinal)) return @"\\?\UNC\" + path.Substring(2);
            return @"\\?\" + path;
        }

        // opt in only - lets an admin list folders the acl would otherwise keep them out of
        public static bool EnableBackupPrivilege()
        {
            if (!IsWindows) return false;
            try
            {
                IntPtr token;
                if (!OpenProcessToken(GetCurrentProcess(), 0x20 | 0x8, out token)) return false;
                try
                {
                    Luid id;
                    if (!LookupPrivilege(null, "SeBackupPrivilege", out id)) return false;
                    TokenPrivileges tp = new TokenPrivileges();
                    tp.Count = 1;
                    tp.Id = id;
                    tp.Attributes = 2;
                    if (!AdjustTokenPrivileges(token, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero)) return false;
                    return Marshal.GetLastWin32Error() == 0;   // 1300 means the token never had it
                }
                finally { CloseHandle(token); }
            }
            catch { return false; }
        }

        public static bool IsElevated()
        {
            try
            {
                using (WindowsIdentity id = WindowsIdentity.GetCurrent())
                    return new WindowsPrincipal(id).IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch { return false; }
        }
    }

    internal struct FileRec
    {
        public string Name;
        public long Size;
        public long Stamp;
        public short Type;
        public byte Guard;
    }

    // file extensions seen so far. one table per scan, shared by its workers and by any later rescan of one folder
    internal sealed class TypeTable
    {
        readonly Dictionary<string, int> ids = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        readonly List<string> list = new List<string>();

        public TypeTable()
        {
            list.Add("");
            ids[""] = 0;
        }

        public int Id(string ext)
        {
            lock (ids)
            {
                int id;
                if (ids.TryGetValue(ext, out id)) return id;
                if (list.Count >= 30000) return 0;     // someone is using random extensions, stop counting them
                id = list.Count;
                list.Add(ext.ToLowerInvariant());
                ids[ext] = id;
                return id;
            }
        }

        public string[] Names()
        {
            lock (ids) return list.ToArray();
        }
    }

    // walks the tree on a crew of threads. the ui only reads the counters til Done flips
    internal sealed class Scanner
    {
        public string RootPath;
        public string GuardRoot;            // where the first scan started, when this one only redoes a folder inside it
        public TypeTable Types;
        public int PerFolder = 5000;        // files that get their own row per folder, the rest become a pile
        public int SoftCap = 2500000;       // past this many file rows we get stingy so ram does not go rubar
        public int Threads;                 // 0 means pick from the cpu count
        public bool Classic;                // list folders the findfirstfile way like the first version did

        public long FilesSeen, DirsSeen, BytesSeen;
        public volatile string Where = "";
        public volatile bool Cancel;
        public volatile bool Done;
        public volatile string Error;

        public Node Root;
        public bool Stopped;
        public bool Capped;
        public int Links;
        public int UnreadableCount;
        public int Crew;
        public double Seconds;
        public readonly List<string> Unreadable = new List<string>();

        readonly object gate = new object();
        readonly Stack<Node> todo = new Stack<Node>();
        int busy, idle;
        long kept;

        static readonly Comparison<FileRec> BigFirst = delegate(FileRec a, FileRec b) { return b.Size.CompareTo(a.Size); };

        public void Start()
        {
            Thread t = new Thread(Run);
            t.IsBackground = true;
            t.Name = "ooutsize scan";
            t.Start();
        }

        void Run()
        {
            Stopwatch sw = Stopwatch.StartNew();
            try
            {
                if (Types == null) Types = new TypeTable();
                Root = new Node();
                Root.Name = RootPath;
                Root.Kind = Node.KDir;
                Root.Type = -1;
                Root.Guard = Guard.Whole(Guard.Below(RootPath, GuardRoot ?? RootPath));
                todo.Push(Root);

                // most of the wait is the disk and the filter drivers sat on top of it, not this code,
                // so going wider than the core count sti;l buys time
                int n = Threads > 0 ? Threads : Math.Max(4, Math.Min(32, Environment.ProcessorCount * 2));
                Crew = n;
                Thread[] threads = new Thread[n];
                for (int i = 0; i < n; i++)
                {
                    threads[i] = new Thread(new Worker(this).Loop);
                    threads[i].IsBackground = true;
                    threads[i].Name = "ooutsize walk " + i;
                    threads[i].Start();
                }
                for (int i = 0; i < n; i++) threads[i].Join();
                Stopped = Cancel;
                Where = "adding it all up";
                RollUp();
            }
            catch (Exception ex) { Error = ex.Message; }
            Seconds = sw.Elapsed.TotalSeconds;
            Done = true;
        }

        // list every folder with parents ahead of their kids, then add up walking that list backwards
        // so a folder is always done after everything inside it - no recursion needed
        void RollUp()
        {
            Comparison<Node> order = new Comparison<Node>(Tree.BigFirst);
            List<Node> all = new List<Node>();
            all.Add(Root);
            for (int i = 0; i < all.Count; i++)
            {
                Node[] kids = all[i].Kids;
                if (kids == null) continue;
                for (int k = 0; k < kids.Length; k++) if (kids[k].Kind == Node.KDir) all.Add(kids[k]);
            }
            for (int i = all.Count - 1; i >= 0; i--) Tree.AddUp(all[i], order);
        }

        // what each file type adds up to, read off the finished tree so it still holds after rows get
        // removed or a folder is scanned again. a pile only counts when everything in it is one type
        public static void Tally(Node top, long[] bytes)
        {
            List<Node> stack = new List<Node>();
            stack.Add(top);
            while (stack.Count > 0)
            {
                Node n = stack[stack.Count - 1];
                stack.RemoveAt(stack.Count - 1);
                if (n.Kind == Node.KDir)
                {
                    Node[] kids = n.Kids;
                    if (kids != null) for (int i = 0; i < kids.Length; i++) stack.Add(kids[i]);
                    continue;
                }
                if (n.Type >= 0 && n.Type < bytes.Length) bytes[n.Type] += n.Size;
            }
        }

        void NoteUnreadable(string path, int code)
        {
            Interlocked.Increment(ref UnreadableCount);
            lock (Unreadable)
            {
                if (Unreadable.Count < 5000) Unreadable.Add(code == 5 || code == 0 ? path : path + "   (error " + code + ")");
            }
        }

        sealed class Worker
        {
            readonly Scanner s;
            readonly List<FileRec> files = new List<FileRec>();
            readonly List<Node> dirs = new List<Node>();
            readonly Dictionary<string, int> seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            readonly byte[] slab;
            readonly string guardRoot;
            char[] letters = new char[512];
            string below = "";
            bool shallow;
            byte inherited;

            public Worker(Scanner owner)
            {
                s = owner;
                guardRoot = owner.GuardRoot ?? owner.RootPath;
                slab = Native.IsWindows && !owner.Classic ? new byte[65536] : null;
            }

            public void Loop()
            {
                try
                {
                    while (true)
                    {
                        Node dir;
                        lock (s.gate)
                        {
                            while (s.todo.Count == 0)
                            {
                                if (s.busy == 0 || s.Cancel) { Monitor.PulseAll(s.gate); return; }
                                s.idle++;
                                Monitor.Wait(s.gate, 200);
                                s.idle--;
                            }
                            if (s.Cancel) { Monitor.PulseAll(s.gate); return; }
                            dir = s.todo.Pop();
                            s.busy++;
                        }
                        try { Folder(dir); }
                        catch (OutOfMemoryException)
                        {
                            s.Error = "Ran out of memory part way through. Try a smaller folder or a lower -MaxFilesPerFolder.";
                            s.Cancel = true;
                        }
                        catch (Exception)
                        {
                            // one bad folder should not sink the who;e scan
                            if (dir.Kids == null) dir.Kids = Tree.None;
                            dir.Denied = true;
                            s.NoteUnreadable(Tree.PathOf(dir), 0);
                        }
                        finally
                        {
                            // only the last one out with nothing left to do needs to wake the others, so they can go home
                            lock (s.gate) { s.busy--; if (s.busy == 0 && s.todo.Count == 0) Monitor.PulseAll(s.gate); }
                        }
                    }
                }
                catch (Exception ex) { s.Error = ex.Message; s.Cancel = true; }
            }

            void Folder(Node dir)
            {
                string path = Tree.PathOf(dir);
                s.Where = path;
                files.Clear();
                dirs.Clear();
                below = Guard.Below(path, guardRoot);
                shallow = Guard.Shallow(below);
                inherited = dir.Guard;
                bool ok;
                if (!Native.IsWindows) ok = ListManaged(dir, path);
                else if (slab != null && ListFast(dir, path)) ok = true;
                else
                {
                    files.Clear();
                    dirs.Clear();
                    ok = ListNative(dir, path);
                }
                for (int i = 0; i < dirs.Count; i++)
                    dirs[i].Guard = shallow ? Guard.Child(below, dirs[i].Name, inherited) : inherited;
                Interlocked.Increment(ref s.DirsSeen);
                if (!ok) { dir.Kids = Tree.None; return; }

                long bytes = 0;
                for (int i = 0; i < files.Count; i++) bytes += files[i].Size;
                Interlocked.Add(ref s.FilesSeen, files.Count);
                Interlocked.Add(ref s.BytesSeen, bytes);

                int keep = files.Count;
                int cap = s.PerFolder;
                if (Interlocked.Read(ref s.kept) > s.SoftCap && (cap <= 0 || cap > 25)) cap = 25;
                if (cap > 0 && keep > cap)
                {
                    files.Sort(BigFirst);
                    keep = cap;
                    if (cap == 25 && s.PerFolder != 25) s.Capped = true;
                }
                int rest = files.Count - keep;

                Node[] kids = new Node[dirs.Count + keep + (rest > 0 ? 1 : 0)];
                int k = 0;
                for (int i = 0; i < dirs.Count; i++) kids[k++] = dirs[i];
                for (int i = 0; i < keep; i++)
                {
                    FileRec f = files[i];
                    Node n = new Node();
                    n.Name = f.Name;
                    n.Parent = dir;
                    n.Kind = Node.KFile;
                    n.Size = f.Size;
                    n.Stamp = f.Stamp;
                    n.Files = 1;
                    n.Type = f.Type;
                    n.Guard = f.Guard;
                    kids[k++] = n;
                }
                if (rest > 0)
                {
                    Node pile = new Node();
                    pile.Parent = dir;
                    pile.Kind = Node.KPile;
                    pile.Files = rest;
                    pile.Guard = inherited;
                    // a pile that is all one type keeps that type, so it gets the colour and counts in the key
                    short same = files[keep].Type;
                    for (int i = keep; i < files.Count; i++)
                    {
                        pile.Size += files[i].Size;
                        if (files[i].Stamp > pile.Stamp) pile.Stamp = files[i].Stamp;
                        if (files[i].Type != same) same = -1;
                    }
                    pile.Type = same;
                    pile.Name = Fmt.Count(rest) + " smaller files";
                    kids[k++] = pile;
                }
                dir.Kids = kids;
                Interlocked.Add(ref s.kept, keep);

                if (dirs.Count > 0)
                {
                    lock (s.gate)
                    {
                        for (int i = 0; i < dirs.Count; i++) s.todo.Push(dirs[i]);
                        // this thread takes one of them itself. wake a sleeper for each of the rest and no more -
                        // waking the whole crew on every push was a stampede that got worse with each thread added
                        int wake = Math.Min(dirs.Count - 1, s.idle);
                        for (int i = 0; i < wake; i++) Monitor.Pulse(s.gate);
                    }
                }
            }

            // one open, then the whole listing comes back in 64k slabs with the size on disk already in each row,
            // so no file has to be opened to ask what it really takes up. findfirstfile and newer .net sit on this
            // same call. false means nothing was listed here and the old lister should have a go
            bool ListFast(Node dir, string path)
            {
                IntPtr h = Native.OpenDir(Native.Long(path));
                if (h == Native.Bad)
                {
                    int err = Marshal.GetLastWin32Error();
                    return err == 2 || err == 3;        // gone since we saw it, same as empty
                }
                try
                {
                    bool first = true;
                    while (true)
                    {
                        Native.IoStatus io;
                        int status = Native.NtQueryDirectoryFile(h, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, out io,
                            slab, slab.Length, 2, false, IntPtr.Zero, first);
                        if (status != 0)
                        {
                            // no more files is the normal way out. an empty volume root says no such file instead
                            if (status == Native.NO_MORE_FILES || status == Native.NO_SUCH_FILE) return true;
                            return !first;
                        }
                        first = false;
                        long got = io.Information.ToInt64();
                        if (got <= 0) return true;
                        // FILE_FULL_DIR_INFORMATION rows chained by the offset in their first four bytes
                        int off = 0;
                        while (off + 68 <= got)
                        {
                            int next = BitConverter.ToInt32(slab, off);
                            int nameBytes = BitConverter.ToInt32(slab, off + 60);
                            if (nameBytes < 0 || off + 68 + nameBytes > got) break;
                            Row(dir, off, nameBytes >> 1);
                            if (next <= 0) break;
                            off += next;
                        }
                        if (s.Cancel) return true;
                    }
                }
                finally { Native.CloseHandle(h); }
            }

            // row layout: write time at 24, size at 40, size on disk at 48, attributes at 56, reparse tag at 64, name at 68
            void Row(Node dir, int off, int chars)
            {
                uint attr = BitConverter.ToUInt32(slab, off + 56);
                if ((attr & Native.DIRECTORY) != 0)
                {
                    if (chars == 1 && slab[off + 68] == 46 && slab[off + 69] == 0) return;
                    if (chars == 2 && slab[off + 68] == 46 && slab[off + 69] == 0 && slab[off + 70] == 46 && slab[off + 71] == 0) return;
                    // junctions and symlinks point somewhere e;se, following them double counts or loops
                    if ((attr & Native.REPARSE) != 0 && (BitConverter.ToUInt32(slab, off + 64) & Native.SURROGATE) != 0)
                    {
                        Interlocked.Increment(ref s.Links);
                        return;
                    }
                    Node c = new Node();
                    c.Name = NameAt(off + 68, chars);
                    c.Parent = dir;
                    c.Kind = Node.KDir;
                    c.Type = -1;
                    dirs.Add(c);
                    return;
                }
                long size = BitConverter.ToInt64(slab, off + 40);
                // sparse, compressed and cloud placeholder files lie about their size, so they count as what is on disk
                if ((attr & Native.THIN) != 0)
                {
                    long onDisk = BitConverter.ToInt64(slab, off + 48);
                    if (onDisk >= 0 && onDisk < size) size = onDisk;
                }
                Take(NameAt(off + 68, chars), size, BitConverter.ToInt64(slab, off + 24));
            }

            // straight copy of the utf-16 so a name with a broken surrogate in it comes through untouched
            string NameAt(int at, int chars)
            {
                if (chars > letters.Length) letters = new char[Math.Max(chars, letters.Length * 2)];
                Buffer.BlockCopy(slab, at, letters, 0, chars * 2);
                return new string(letters, 0, chars);
            }

            bool ListNative(Node dir, string path)
            {
                string longPath = Native.Long(path);
                bool slash = longPath.EndsWith("\\", StringComparison.Ordinal);
                Native.FindData fd;
                // info level 1 skips the 8.3 name, flag 2 asks for bigger batches - both speed things up
                IntPtr h = Native.FindFirst(longPath + (slash ? "*" : "\\*"), 1, out fd, 0, IntPtr.Zero, 2);
                if (h == Native.Bad)
                {
                    int err = Marshal.GetLastWin32Error();
                    if (err == 2 || err == 3 || err == 18) return true;     // empty, or gone since we saw it
                    dir.Denied = true;
                    s.NoteUnreadable(path, err);
                    return false;
                }
                try
                {
                    int n = 0;
                    do
                    {
                        if ((++n & 255) == 0 && s.Cancel) break;
                        string name = fd.Name;
                        uint attr = fd.Attr;
                        if ((attr & Native.DIRECTORY) != 0)
                        {
                            if (name == "." || name == "..") continue;
                            if ((attr & Native.REPARSE) != 0 && (fd.Tag & Native.SURROGATE) != 0)
                            {
                                Interlocked.Increment(ref s.Links);
                                continue;
                            }
                            Node c = new Node();
                            c.Name = name;
                            c.Parent = dir;
                            c.Kind = Node.KDir;
                            c.Type = -1;
                            dirs.Add(c);
                        }
                        else
                        {
                            long size = ((long)fd.SizeHi << 32) | fd.SizeLo;
                            // this lister has no size on disk in its rows so thin files cost an extra call each
                            if ((attr & Native.THIN) != 0)
                            {
                                uint hi;
                                uint lo = Native.CompressedSize(longPath + (slash ? "" : "\\") + name, out hi);
                                if (lo != 0xFFFFFFFF || Marshal.GetLastWin32Error() == 0)
                                {
                                    long onDisk = ((long)hi << 32) | lo;
                                    if (onDisk < size) size = onDisk;
                                }
                            }
                            Take(name, size, ((long)fd.WriteHi << 32) | fd.WriteLo);
                        }
                    }
                    while (Native.FindNext(h, out fd));
                }
                finally { Native.FindClose(h); }
                return true;
            }

            // only used off windows, which is how this gets exercised on a linux box
            bool ListManaged(Node dir, string path)
            {
                try
                {
                    int n = 0;
                    foreach (FileSystemInfo e in new DirectoryInfo(path).EnumerateFileSystemInfos())
                    {
                        if ((++n & 255) == 0 && s.Cancel) break;
                        try
                        {
                            FileAttributes attr = e.Attributes;
                            if ((attr & FileAttributes.Directory) != 0)
                            {
                                if ((attr & FileAttributes.ReparsePoint) != 0) { Interlocked.Increment(ref s.Links); continue; }
                                Node c = new Node();
                                c.Name = e.Name;
                                c.Parent = dir;
                                c.Kind = Node.KDir;
                                c.Type = -1;
                                dirs.Add(c);
                            }
                            else
                            {
                                FileInfo fi = e as FileInfo;
                                Take(e.Name, fi == null ? 0 : fi.Length, e.LastWriteTimeUtc.ToFileTimeUtc());
                            }
                        }
                        catch (IOException) { }
                        catch (UnauthorizedAccessException) { }
                    }
                    return true;
                }
                catch (UnauthorizedAccessException) { dir.Denied = true; s.NoteUnreadable(path, 5); return false; }
                catch (DirectoryNotFoundException) { return true; }
                catch (IOException) { dir.Denied = true; s.NoteUnreadable(path, 0); return false; }
                catch (System.Security.SecurityException) { dir.Denied = true; s.NoteUnreadable(path, 5); return false; }
            }

            void Take(string name, long size, long stamp)
            {
                FileRec f;
                f.Guard = shallow ? Guard.Child(below, name, inherited) : inherited;
                f.Name = name;
                f.Size = size;
                f.Stamp = stamp;
                f.Type = (short)TypeOf(name);
                files.Add(f);
            }

            int TypeOf(string name)
            {
                int dot = name.LastIndexOf('.');
                string ext = "";
                if (dot > 0 && dot < name.Length - 1 && name.Length - dot <= 11)
                {
                    ext = name.Substring(dot);
                    if (ext.IndexOf(' ') >= 0) ext = "";
                }
                int id;
                if (!seen.TryGetValue(ext, out id))
                {
                    id = s.Types.Id(ext);
                    seen[ext] = id;
                }
                return id;
            }
        }
    }

    // colours, fonts and the drawing helpers everything shares. all of it follows the gbl style guide
    internal static class Ui
    {
        // the guide's named colours and nothing mixed by eye. neutrals first, then each hue as its colour,
        // its light background and (where one gets used here) its dark step for text
        public static readonly Color White = Rgb(0xFFFFFF), GrayLight = Rgb(0xE8ECF1), GrayLine = Rgb(0xDDE3EA),
            GrayBorder = Rgb(0xC9D2DC), Gray = Rgb(0x5B6773), Graphite = Rgb(0x1F2730), Navy = Rgb(0x16304F),
            Yellow = Rgb(0xF5C400), YellowLight = Rgb(0xFCF5CB),
            Blue = Rgb(0x2F6FD6), BlueLight = Rgb(0xE4ECF9),
            Green = Rgb(0x3B8132), GreenLight = Rgb(0xACD8A7),
            Aqua = Rgb(0x17A9CB), AquaLight = Rgb(0xDFF0F6),
            Orange = Rgb(0xEE7A1A), OrangeLight = Rgb(0xFCEADB), OrangeText = Rgb(0xAD4A00),
            Purple = Rgb(0x7B4FA6), PurpleLight = Rgb(0xEEE8F5),
            Red = Rgb(0xB5443A), RedLight = Rgb(0xFAE6E3), RedText = Rgb(0xA7372F),
            Rose = Rgb(0xE573A3), RoseLight = Rgb(0xF7E6EC);

        // the guide dims a mark to 30 percent when another one is picked out. worked out against white once so a
        // dimmed block is a plain fill. everything that is not one of the coloured file types is drawn that way for good
        public static readonly Color Other = Faded(Orange);
        public static readonly Color Pick = BlueLight, PickIdle = Mix(BlueLight, White, 0.45), Hot = Mix(BlueLight, White, 0.65);

        // chart colours in the guide's order of preference, biggest file type first. red sits out, it is for status only
        public static readonly Color[] TypeFill = { Yellow, Blue, Green, Aqua, Orange, Purple, Rose };
        public static readonly Color[] TypeDim = { Faded(Yellow), Faded(Blue), Faded(Green), Faded(Aqua), Faded(Orange),
            Faded(Purple), Faded(Rose) };

        // tiles in that same order - light background, bar in its own colour. there is no red tile
        public static readonly Color[] TileLight = { YellowLight, BlueLight, GreenLight, AquaLight, OrangeLight, PurpleLight, RoseLight };
        public static readonly Color[] TileBar = { Yellow, Blue, Green, Aqua, Orange, Purple, Rose };

        // arial throughout, regular and bold only, at the guide's pixel sizes: body 13, label 11.5, caption 10, heading 15 bold.
        // the tile numeral is smaller than the guide's 40 because these tiles are
        public static readonly Font Base = Make(13f, false), Bold = Make(13f, true), Label = Make(11.5f, false),
            LabelBold = Make(11.5f, true), Caption = Make(10f, false), Head = Make(15f, true), Numeral = Make(22f, false);

        public static readonly StringFormat Left = Format(StringAlignment.Near), Right = Format(StringAlignment.Far),
            Mid = Format(StringAlignment.Center), Wrap = new StringFormat();

        static readonly Dictionary<int, SolidBrush> brushes = new Dictionary<int, SolidBrush>();
        static readonly Dictionary<int, Pen> pens = new Dictionary<int, Pen>();
        static readonly Dictionary<int, Color> inks = new Dictionary<int, Color>();

        static Color Rgb(int v) { return Color.FromArgb(255, (v >> 16) & 255, (v >> 8) & 255, v & 255); }

        // t is how far to go from a to b
        public static Color Mix(Color a, Color b, double t)
        {
            return Color.FromArgb(255, (int)Math.Round(a.R + (b.R - a.R) * t), (int)Math.Round(a.G + (b.G - a.G) * t),
                (int)Math.Round(a.B + (b.B - a.B) * t));
        }

        public static Color Faded(Color c) { return Mix(c, White, 0.70); }

        static Font Make(float pixels, bool bold)
        {
            FontStyle style = bold ? FontStyle.Bold : FontStyle.Regular;
            try { return new Font("Arial", pixels, style, GraphicsUnit.Pixel); }
            catch { return new Font(FontFamily.GenericSansSerif, pixels, style, GraphicsUnit.Pixel); }
        }

        static StringFormat Format(StringAlignment a)
        {
            StringFormat f = new StringFormat(StringFormatFlags.NoWrap);
            f.Alignment = a;
            f.LineAlignment = StringAlignment.Center;
            f.Trimming = StringTrimming.EllipsisCharacter;
            return f;
        }

        public static SolidBrush B(Color c)
        {
            SolidBrush b;
            int key = c.ToArgb();
            if (!brushes.TryGetValue(key, out b)) { b = new SolidBrush(c); brushes[key] = b; }
            return b;
        }

        public static Pen P(Color c)
        {
            Pen p;
            int key = c.ToArgb();
            if (!pens.TryGetValue(key, out p)) { p = new Pen(c); pens[key] = p; }
            return p;
        }

        public static void Text(Graphics g, string s, Font f, Color c, float x, float y, float w, float h, StringFormat sf)
        {
            if (w < 4 || h < 4 || string.IsNullOrEmpty(s)) return;
            g.DrawString(s, f, B(c), new RectangleF(x, y, w, h), sf);
        }

        public static int Width(Graphics g, string s, Font f)
        {
            return (int)Math.Ceiling(g.MeasureString(s, f).Width);
        }

        static double Channel(int v)
        {
            double s = v / 255.0;
            return s <= 0.03928 ? s / 12.92 : Math.Pow((s + 0.055) / 1.055, 2.4);
        }

        static double Luminance(Color c) { return 0.2126 * Channel(c.R) + 0.7152 * Channel(c.G) + 0.0722 * Channel(c.B); }

        // text that sits on a coloured block is graphite or white, whichever of the two reads better against the fill
        public static Color InkOn(Color fill)
        {
            Color ink;
            int key = fill.ToArgb();
            if (inks.TryGetValue(key, out ink)) return ink;
            double l = Luminance(fill);
            ink = (l + 0.05) / (Luminance(Graphite) + 0.05) >= 1.05 / (l + 0.05) ? Graphite : White;
            inks[key] = ink;
            return ink;
        }

        // 0 to 1 in, 0 to 1 out, quick off the mark and gentle at the end. close enough to the guide's curve
        public static float Ease(float t)
        {
            float u = 1f - t;
            return 1f - u * u * u;
        }

        // rounded outline with a radius per corner, clockwise from top left. 0 leaves that corner square
        public static GraphicsPath Round(float x, float y, float w, float h, float tl, float tr, float br, float bl)
        {
            float most = Math.Min(w, h) / 2f;
            tl = Math.Min(tl, most);
            tr = Math.Min(tr, most);
            br = Math.Min(br, most);
            bl = Math.Min(bl, most);
            GraphicsPath p = new GraphicsPath();
            if (tl > 0) p.AddArc(x, y, tl * 2, tl * 2, 180, 90); else p.AddLine(x, y, x, y);
            if (tr > 0) p.AddArc(x + w - tr * 2, y, tr * 2, tr * 2, 270, 90); else p.AddLine(x + w, y, x + w, y);
            if (br > 0) p.AddArc(x + w - br * 2, y + h - br * 2, br * 2, br * 2, 0, 90); else p.AddLine(x + w, y + h, x + w, y + h);
            if (bl > 0) p.AddArc(x, y + h - bl * 2, bl * 2, bl * 2, 90, 90); else p.AddLine(x, y + h, x, y + h);
            p.CloseFigure();
            return p;
        }

        // rounded box. an edge is the box filled twice, once in the edge colour and once a pixel in,
        // which stays crisp where a pen would smear. pass Color.Empty for no edge
        public static void Box(Graphics g, float x, float y, float w, float h, float r, Color fill, Color edge)
        {
            if (w < 1 || h < 1) return;
            SmoothingMode sm = g.SmoothingMode;
            PixelOffsetMode pm = g.PixelOffsetMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.Half;
            using (GraphicsPath outer = Round(x, y, w, h, r, r, r, r)) g.FillPath(B(edge.IsEmpty ? fill : edge), outer);
            if (!edge.IsEmpty && w > 2 && h > 2)
            {
                float q = Math.Max(0f, r - 1f);
                using (GraphicsPath inner = Round(x + 1, y + 1, w - 2, h - 2, q, q, q, q)) g.FillPath(B(fill), inner);
            }
            g.SmoothingMode = sm;
            g.PixelOffsetMode = pm;
        }

        // the guide's tile - light background, 3px bar across the top in the tile's own colour, 6px corners
        public static void Tile(Graphics g, float x, float y, float w, float h, Color light, Color bar)
        {
            Box(g, x, y, w, h, 6f, light, Color.Empty);
            TopBar(g, x, y, w, Math.Min(6f, Math.Min(w, h) / 2f), bar);
        }

        // the guide's card - white, gray-line edge, 3px orange bar on top, 6px corners. the shadow is up to the caller
        public static void Card(Graphics g, float x, float y, float w, float h)
        {
            Box(g, x, y, w, h, 6f, White, GrayLine);
            TopBar(g, x, y, w, Math.Min(6f, Math.Min(w, h) / 2f), Orange);
        }

        // the 3px bar. it follows the rounded top corners the way a css border on a rounded box does
        static void TopBar(Graphics g, float x, float y, float w, float r, Color bar)
        {
            if (r <= 3f) { g.FillRectangle(B(bar), x, y, w, 3f); return; }
            SmoothingMode sm = g.SmoothingMode;
            PixelOffsetMode pm = g.PixelOffsetMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.Half;
            // where the corner's curve is 3px down from the top, as an angle on its circle
            float a = 180f + (float)(Math.Asin((r - 3f) / r) * 180.0 / Math.PI);
            using (GraphicsPath top = new GraphicsPath())
            {
                top.AddArc(x, y, r * 2, r * 2, a, 270f - a);
                top.AddArc(x + w - r * 2, y, r * 2, r * 2, 270f, 270f - a);
                top.CloseFigure();
                g.FillPath(B(bar), top);
            }
            g.SmoothingMode = sm;
            g.PixelOffsetMode = pm;
        }

        // soft shadow for something lifted off the page. gdi has no blur, so it is rings of graphite, each a
        // pixel wider than the last and all of them faint. alpha is about what the middle adds up to, out of 255
        public static void Shadow(Graphics g, float x, float y, float w, float h, float r, float drop, int spread, int alpha)
        {
            if (alpha <= 0 || spread <= 0) return;
            Brush ring = B(Color.FromArgb(Math.Max(1, alpha / spread), Graphite));
            SmoothingMode sm = g.SmoothingMode;
            PixelOffsetMode pm = g.PixelOffsetMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.Half;
            for (int i = spread; i >= 1; i--)
                using (GraphicsPath p = Round(x - i, y + drop - i, w + 2 * i, h + 2 * i, r + i, r + i, r + i, r + i)) g.FillPath(ring, p);
            g.SmoothingMode = sm;
            g.PixelOffsetMode = pm;
        }

        // a meter - square where it starts, round where it ends. the track behind it is the pale neutral,
        // so the only colour is the part that measures something
        public static void Bar(Graphics g, int x, int y, int w, int h, double share, Color fill, Color track)
        {
            Pill(g, B(track), x, y, w, h);
            if (share <= 0) return;
            int fw = (int)Math.Round(w * Math.Min(1.0, share));
            if (fw < 2) fw = 2;
            Pill(g, B(fill), x, y, fw, h);
        }

        static void Pill(Graphics g, Brush b, int x, int y, int w, int h)
        {
            if (w <= h) { g.FillRectangle(b, x, y, w, h); return; }
            g.FillRectangle(b, x, y, w - h / 2, h);
            SmoothingMode sm = g.SmoothingMode;
            PixelOffsetMode pm = g.PixelOffsetMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.Half;
            g.FillPie(b, x + w - h, y, h, h, -90, 180);
            g.SmoothingMode = sm;
            g.PixelOffsetMode = pm;
        }
    }

    // the left pane. home made tree with columns because winforms has no such thing
    // and owner drawing the stock treeview is a sht show
    internal sealed class TreeGrid : Control
    {
        const int RowH = 22, HeadH = 24, Step = 16;
        Node root, sel;
        readonly List<Node> rows = new List<Node>();
        readonly List<int> depth = new List<int>();
        int top, hot = -1;
        int deepest;             // deepest level showing right now
        bool quiet;
        readonly VScrollBar bar = new VScrollBar();

        public sbyte[] Slot;
        public event NodeEvent Picked;
        public event NodeEvent Opened;
        public event NodeAtEvent MenuWanted;

        public TreeGrid()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
            TabStop = true;
            BackColor = Ui.White;
            bar.Visible = false;
            bar.ValueChanged += delegate { if (!quiet) { top = bar.Value; Invalidate(); } };
            Controls.Add(bar);
        }

        public Node Selected { get { return sel; } }

        public void SetRoot(Node r)
        {
            root = r;
            sel = r;
            top = 0;
            hot = -1;
            if (r != null) r.Open = true;
            Rebuild();
        }

        // rows were taken out or a folder was swapped for a fresh copy - redo the list and stay near where we were
        public void Reload(Node keep)
        {
            if (root == null) return;
            hot = -1;
            sel = keep != null && Tree.Under(keep, root) ? keep : root;
            Rebuild();
            Select(sel, false);
        }

        void Rebuild()
        {
            rows.Clear();
            depth.Clear();
            deepest = 0;
            if (root != null) AddRows(root, 0);
            Fit();
            Invalidate();
        }

        void AddRows(Node n, int d)
        {
            rows.Add(n);
            depth.Add(d);
            if (d > deepest) deepest = d;
            if (n.Open && n.Kids != null)
                for (int i = 0; i < n.Kids.Length; i++) AddRows(n.Kids[i], d + 1);
        }

        int Page { get { return Math.Max(1, (ClientSize.Height - HeadH) / RowH); } }

        void Fit()
        {
            int page = Page;
            int maxTop = Math.Max(0, rows.Count - page);
            if (top > maxTop) top = maxTop;
            if (top < 0) top = 0;
            quiet = true;
            try
            {
                int bw = SystemInformation.VerticalScrollBarWidth;
                bar.SetBounds(ClientSize.Width - bw, HeadH, bw, Math.Max(0, ClientSize.Height - HeadH));
                bool need = rows.Count > page;
                if (need)
                {
                    bar.Minimum = 0;
                    bar.Maximum = rows.Count - 1;
                    bar.SmallChange = 1;
                    bar.LargeChange = page;
                    bar.Value = Math.Min(top, bar.Maximum);
                }
                bar.Visible = need;
            }
            finally { quiet = false; }
        }

        void ScrollTo(int row)
        {
            top = row;
            Fit();
            Invalidate();
        }

        public void Select(Node n, bool byUser)
        {
            if (n == null || root == null || !Tree.Under(n, root)) return;
            bool opened = false;
            for (Node p = n.Parent; p != null; p = p.Parent)
                if (!p.Open) { p.Open = true; opened = true; }
            if (opened) { rows.Clear(); depth.Clear(); deepest = 0; AddRows(root, 0); }
            sel = n;
            int i = rows.IndexOf(n);
            if (i >= 0)
            {
                int page = Page;
                if (i < top) top = byUser ? i : Math.Max(0, i - page / 3);
                else if (i >= top + page) top = byUser ? i - page + 1 : Math.Max(0, i - page / 3);
            }
            Fit();
            Invalidate();
            if (byUser && Picked != null) Picked(n);
        }

        void Toggle(Node n)
        {
            if (!Tree.HasKids(n)) return;
            n.Open = !n.Open;
            bool moved = false;
            if (!n.Open && sel != null && sel != n && Tree.Under(sel, n)) { sel = n; moved = true; }
            Rebuild();
            if (moved && Picked != null) Picked(n);
        }

        void Columns(out int w, out int xSize, out int xPct, out int xBar, out int xFiles, out int xDate)
        {
            w = ClientSize.Width - (bar.Visible ? bar.Width : 0);
            int x = w - 8;
            xDate = -1;
            xFiles = -1;
            // the name column keeps about 250 px before the optional columns are allowed in
            if (w >= 548) { x -= 82; xDate = x; }
            if (w >= 614) { x -= 66; xFiles = x; }
            x -= 84; xBar = x;
            x -= 50; xPct = x;
            x -= 74; xSize = x;
        }

        int Indent(int row, int xSize)
        {
            // deep paths would shove the name off the edge, so the step per level shrinks as the tree gets deeper
            int step = Math.Min(Step, Math.Max(5, (xSize - 170) / Math.Max(1, deepest)));
            return 6 + depth[row] * step;
        }

        static double Share(Node n)
        {
            if (n.Parent == null) return 1;
            return n.Parent.Size > 0 ? (double)n.Size / n.Parent.Size : 0;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(Ui.White);
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            int w, xSize, xPct, xBar, xFiles, xDate;
            Columns(out w, out xSize, out xPct, out xBar, out xFiles, out xDate);

            // text is graphite, the quieter columns and the headings are the guide's grey, lines are its hairline
            g.DrawLine(Ui.P(Ui.GrayBorder), 0, HeadH - 1, ClientSize.Width, HeadH - 1);
            long stale = DateTime.UtcNow.AddDays(-365).ToFileTimeUtc();
            Ui.Text(g, "Name", Ui.Label, Ui.Gray, 10, 0, xSize - 16, HeadH, Ui.Left);
            Ui.Text(g, "Size", Ui.Label, Ui.Gray, xSize, 0, 70, HeadH, Ui.Right);
            Ui.Text(g, "Share of parent", Ui.Label, Ui.Gray, xPct, 0, xBar + 80 - xPct, HeadH, Ui.Right);
            if (xFiles >= 0) Ui.Text(g, "Files", Ui.Label, Ui.Gray, xFiles, 0, 62, HeadH, Ui.Right);
            if (xDate >= 0) Ui.Text(g, "Newest", Ui.Label, Ui.Gray, xDate, 0, 78, HeadH, Ui.Right);

            for (int i = 0; ; i++)
            {
                int r = top + i;
                int y = HeadH + i * RowH;
                if (r >= rows.Count || y >= ClientSize.Height) break;
                Node n = rows[r];
                bool picked = n == sel;
                if (picked) g.FillRectangle(Ui.B(Focused ? Ui.Pick : Ui.PickIdle), 0, y, w, RowH);
                else if (r == hot) g.FillRectangle(Ui.B(Ui.Hot), 0, y, w, RowH);

                int x = Indent(r, xSize);
                if (Tree.HasKids(n)) Twisty(g, x, y, n.Open);
                x += 14;
                Glyph(g, n, x, y + 5);
                x += 18;
                string name = n.Denied ? n.Name + "   (could not be read)" : n.Name;
                Ui.Text(g, name, Ui.Base, n.Kind == Node.KPile ? Ui.Gray : Ui.Graphite, x, y, xSize - 6 - x, RowH, Ui.Left);
                Ui.Text(g, Fmt.Size(n.Size), Ui.Base, Ui.Graphite, xSize, y, 70, RowH, Ui.Right);
                double share = Share(n);
                Ui.Text(g, Fmt.Pct(share), Ui.Base, Ui.Gray, xPct, y, 46, RowH, Ui.Right);
                // one series of bars, so it takes the first colour in the guide's order
                Ui.Bar(g, xBar + 8, y + 7, 72, 8, share, Ui.Yellow, Ui.GrayLight);
                if (xFiles >= 0 && n.Kind != Node.KFile) Ui.Text(g, Fmt.Count(n.Files), Ui.Base, Ui.Gray, xFiles, y, 62, RowH, Ui.Right);
                if (xDate >= 0 && n.Stamp > 0 && n.Stamp < stale)
                {
                    // nothing in here has changed for over a year. orange is for what wants a look, so it gets the orange tag
                    string when = Fmt.Date(n.Stamp);
                    int tw = Ui.Width(g, when, Ui.Label) + 8;
                    Ui.Box(g, xDate + 78 - tw, y + 3, tw, RowH - 7, 3, Ui.OrangeLight, Color.Empty);
                    Ui.Text(g, when, Ui.Label, Ui.OrangeText, xDate + 78 - tw, y + 3, tw, RowH - 7, Ui.Mid);
                }
                else if (xDate >= 0) Ui.Text(g, Fmt.Date(n.Stamp), Ui.Base, Ui.Gray, xDate, y, 78, RowH, Ui.Right);
                g.DrawLine(Ui.P(Ui.GrayLine), 0, y + RowH - 1, w, y + RowH - 1);
            }
        }

        static void Twisty(Graphics g, int x, int y, bool open)
        {
            SmoothingMode sm = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            Point[] pts = open
                ? new Point[] { new Point(x + 1, y + 9), new Point(x + 11, y + 9), new Point(x + 6, y + 15) }
                : new Point[] { new Point(x + 3, y + 6), new Point(x + 9, y + 11), new Point(x + 3, y + 16) };
            g.FillPolygon(Ui.B(Ui.Gray), pts);
            g.SmoothingMode = sm;
        }

        void Glyph(Graphics g, Node n, int x, int y)
        {
            if (n.Kind == Node.KDir)
            {
                // little folder shape in yellow, orange when it could not be read
                Brush b = Ui.B(n.Denied ? Ui.Orange : Ui.Yellow);
                g.FillRectangle(b, x, y + 1, 6, 3);
                g.FillRectangle(b, x, y + 3, 13, 8);
                return;
            }
            Color c = Ui.Other;
            if (Slot != null && n.Type >= 0 && n.Type < Slot.Length && Slot[n.Type] >= 0)
                c = Ui.TypeFill[Slot[n.Type]];
            if (n.Kind == Node.KPile) g.FillRectangle(Ui.B(Ui.GrayBorder), x + 4, y, 9, 9);
            g.FillRectangle(Ui.B(c), x + 1, y + 2, 9, 9);
        }

        int RowAt(int y)
        {
            if (y < HeadH) return -1;
            int r = top + (y - HeadH) / RowH;
            return r < rows.Count ? r : -1;
        }

        bool OnTwisty(int row, int x)
        {
            int w, xSize, xPct, xBar, xFiles, xDate;
            Columns(out w, out xSize, out xPct, out xBar, out xFiles, out xDate);
            int tx = Indent(row, xSize);
            return x >= tx - 3 && x < tx + 14 && Tree.HasKids(rows[row]);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            Focus();
            int r = RowAt(e.Y);
            if (r < 0) return;
            Node n = rows[r];
            if (e.Button == MouseButtons.Left && OnTwisty(r, e.X)) { Toggle(n); return; }
            if (n != sel) Select(n, true);
            if (e.Button == MouseButtons.Right && MenuWanted != null) MenuWanted(n, PointToScreen(e.Location));
        }

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            base.OnMouseDoubleClick(e);
            int r = RowAt(e.Y);
            if (r < 0 || e.Button != MouseButtons.Left || OnTwisty(r, e.X)) return;
            Node n = rows[r];
            Toggle(n);
            if (Opened != null) Opened(n);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            int r = RowAt(e.Y);
            if (r != hot) { hot = r; Invalidate(); }
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            if (hot != -1) { hot = -1; Invalidate(); }
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            int lines = e.Delta * Math.Max(1, SystemInformation.MouseWheelScrollLines) / 120;
            if (lines == 0) lines = Math.Sign(e.Delta);
            ScrollTo(top - lines);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            Fit();
        }

        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
        protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }

        // winforms fkn eats the arrow keys for focus moves unless we claim them
        protected override bool IsInputKey(Keys keyData)
        {
            Keys k = keyData & Keys.KeyCode;
            if (k == Keys.Up || k == Keys.Down || k == Keys.Left || k == Keys.Right || k == Keys.Enter) return true;
            return base.IsInputKey(keyData);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (rows.Count == 0) return;
            int i = sel == null ? -1 : rows.IndexOf(sel);
            int go = -1;
            switch (e.KeyCode)
            {
                case Keys.Up: go = Math.Max(0, i - 1); break;
                case Keys.Down: go = Math.Min(rows.Count - 1, i + 1); break;
                case Keys.PageUp: go = Math.Max(0, i - Page); break;
                case Keys.PageDown: go = Math.Min(rows.Count - 1, i + Page); break;
                case Keys.Home: go = 0; break;
                case Keys.End: go = rows.Count - 1; break;
                case Keys.Right:
                    if (Tree.HasKids(sel)) { if (!sel.Open) Toggle(sel); else go = Math.Min(rows.Count - 1, i + 1); }
                    break;
                case Keys.Left:
                    if (Tree.HasKids(sel) && sel.Open) Toggle(sel);
                    else if (sel != null && sel != root && sel.Parent != null) go = rows.IndexOf(sel.Parent);
                    break;
                case Keys.Enter:
                    if (sel != null) { Toggle(sel); if (Opened != null) Opened(sel); }
                    break;
                default: return;
            }
            e.Handled = true;
            if (go >= 0 && go < rows.Count && rows[go] != sel) Select(rows[go], true);
        }
    }

    internal struct Hit
    {
        public Node N;
        public int X, Y, W, H;
    }

    internal sealed class DriveRow
    {
        public string Root, Label;
        public long Total, Free;
        public Rectangle Box;
    }

    // the right pane - drive picker when idle, progress while scanning, treemap after
    internal sealed class MapView : Control
    {
        const int Idle = 0, Busy = 1, Map = 2;
        int mode = Idle;
        Node view, pick, hover;
        Point mouse;
        Bitmap bmp;
        bool stale = true;
        readonly List<Hit> hits = new List<Hit>();
        sbyte[] slot;
        int emph = -1;
        int preview = -1;       // a type the mouse is on in the key. it wins over emph while it lasts
        readonly List<DriveRow> drives = new List<DriveRow>();
        int hotDrive = -1;
        float[] rise = new float[0];        // how far each drive card has come up, 0 to 1
        readonly System.Windows.Forms.Timer anim = new System.Windows.Forms.Timer();
        static readonly Pen Dash = MakeDash();

        public string BusyTitle = "", BusyLine1 = "", BusyLine2 = "";
        public double BusyShare = -1;

        public event NodeEvent Picked;
        public event NodeEvent Hovered;
        public event NodeEvent ZoomWanted;
        public event NodeAtEvent MenuWanted;
        public event TextEvent DriveWanted;

        public MapView()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
            BackColor = Ui.White;
            anim.Interval = 15;
            anim.Tick += delegate { Settle(); };
        }

        static Pen MakeDash()
        {
            Pen p = new Pen(Ui.Navy, 1f);
            p.DashPattern = new float[] { 3f, 3f };
            return p;
        }

        // drive cards come up under the mouse and go back down when it leaves, over 0.18s like the guide asks
        void Settle()
        {
            bool moving = false;
            for (int i = 0; i < rise.Length; i++)
            {
                float want = mode == Idle && i == hotDrive ? 1f : 0f;
                if (rise[i] == want) continue;
                rise[i] = rise[i] < want ? Math.Min(want, rise[i] + 15f / 180f) : Math.Max(want, rise[i] - 15f / 180f);
                moving = true;
            }
            if (!moving) anim.Stop();
            Invalidate();
        }

        public void ShowIdle()
        {
            mode = Idle;
            view = pick = hover = null;
            hits.Clear();
            drives.Clear();
            try
            {
                foreach (DriveInfo d in DriveInfo.GetDrives())
                {
                    try
                    {
                        // fixed disks only. poking a dead network drive or an empty cd tray can hang for ages
                        if (d.DriveType != DriveType.Fixed || !d.IsReady) continue;
                        DriveRow r = new DriveRow();
                        r.Root = d.Name;
                        r.Total = d.TotalSize;
                        r.Free = d.TotalFreeSpace;
                        r.Label = d.VolumeLabel;
                        if (r.Total > 0) drives.Add(r);
                    }
                    catch { }
                }
            }
            catch { }
            rise = new float[drives.Count];
            hotDrive = -1;
            Invalidate();
        }

        public void ShowBusy()
        {
            mode = Busy;
            view = pick = hover = null;
            hits.Clear();
            Invalidate();
        }

        public void ShowMap(Node root, sbyte[] typeSlot)
        {
            mode = Map;
            slot = typeSlot;
            emph = -1;
            preview = -1;
            view = root;
            pick = hover = null;
            hits.Clear();
            stale = true;
            Invalidate();
        }

        public void SetView(Node n)
        {
            view = n;
            hover = null;
            hits.Clear();       // the old boxes are for the old view, a mouse move before the repaint must not find them
            stale = true;
            Invalidate();
        }

        public void SetPick(Node n)
        {
            pick = n;
            Invalidate();
        }

        public void SetSlots(sbyte[] typeSlot)
        {
            slot = typeSlot;
            stale = true;
            Invalidate();
        }

        public void SetEmphasis(int s)
        {
            emph = s;
            stale = true;
            Invalidate();
        }

        public void SetPreview(int s)
        {
            if (s == preview) return;
            preview = s;
            stale = true;
            Invalidate();
        }

        // the type that is picked out right now, -1 for none
        int Shown { get { return preview >= 0 ? preview : emph; } }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            hits.Clear();
            hover = null;
            stale = true;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && bmp != null) { bmp.Dispose(); bmp = null; }
            if (disposing) anim.Dispose();
            base.Dispose(disposing);
        }

        // ---------- painting

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            if (mode == Idle) { PaintDrives(g); return; }
            if (mode == Busy) { PaintBusy(g); return; }
            if (stale || bmp == null || bmp.Width != ClientSize.Width || bmp.Height != ClientSize.Height) Render(g.DpiX, g.DpiY);
            if (bmp == null) { g.Clear(Ui.White); return; }
            g.DrawImageUnscaled(bmp, 0, 0);

            Rectangle r;
            if (pick != null && FindBox(pick, out r)) Frame(g, r, true);
            if (hover != null && hover != pick && FindBox(hover, out r)) Frame(g, r, false);
            if (hover != null) Tip(g);
        }

        // navy is the guide's colour for a hover cursor and nothing else, a thin dashed line. what is picked gets
        // the same line solid and doubled. white rides along so either one shows on a dark block too
        static void Frame(Graphics g, Rectangle r, bool picked)
        {
            if (picked)
            {
                g.DrawRectangle(Ui.P(Ui.Navy), r.X, r.Y, r.Width - 1, r.Height - 1);
                g.DrawRectangle(Ui.P(Ui.Navy), r.X + 1, r.Y + 1, r.Width - 3, r.Height - 3);
                g.DrawRectangle(Ui.P(Ui.White), r.X + 2, r.Y + 2, r.Width - 5, r.Height - 5);
                return;
            }
            g.DrawRectangle(Ui.P(Ui.White), r.X, r.Y, r.Width - 1, r.Height - 1);
            g.DrawRectangle(Dash, r.X, r.Y, r.Width - 1, r.Height - 1);
        }

        // the thing itself may be too small to have a box of its own, so walk up til something does
        bool FindBox(Node n, out Rectangle r)
        {
            // a folder that got skipped as a lone wrapper is drawn as the folder inside it
            for (Node q = n; q != null; q = q.Kids != null && q.Kids.Length == 1 ? q.Kids[0] : null)
                if (BoxOf(q, out r)) return true;
            for (Node p = n.Parent; p != null && p != view; p = p.Parent)
                if (BoxOf(p, out r)) return true;
            r = Rectangle.Empty;
            return false;
        }

        bool BoxOf(Node n, out Rectangle r)
        {
            for (int i = hits.Count - 1; i >= 0; i--)
            {
                if (hits[i].N != n) continue;
                r = new Rectangle(hits[i].X, hits[i].Y, hits[i].W, hits[i].H);
                return r.Width > 2 && r.Height > 2;
            }
            r = Rectangle.Empty;
            return false;
        }

        // the guide's tooltip - white, gray-border edge, 4px corners and a small shadow. name in bold on top,
        // then a swatch with the size in bold, then the smaller grey note
        void Tip(Graphics g)
        {
            Node n = hover;
            string size = Fmt.Size(n.Size);
            string share = view != null && view.Size > 0 ? Fmt.Pct((double)n.Size / view.Size) + " of this view" : "";
            string more = n.Kind == Node.KDir ? Fmt.Things(n.Files, "file", "files") : "";
            if (n.Kind == Node.KDir && n.Dirs > 0) more += ", " + Fmt.Things(n.Dirs, "folder", "folders");
            string when = Fmt.Date(n.Stamp);
            if (when.Length > 0) more += (more.Length > 0 ? ", " : "") + (n.Kind == Node.KFile ? "changed " : "newest ") + when;

            int sizeW = Ui.Width(g, size, Ui.LabelBold);
            int w = Math.Max(Ui.Width(g, n.Name, Ui.LabelBold), Math.Max(22 + sizeW + Ui.Width(g, share, Ui.Label), Ui.Width(g, more, Ui.Caption)));
            string advice = Guard.Advice[n.Guard < Guard.Advice.Length ? n.Guard : 0];
            if (advice.Length > 0) w = Math.Max(w, 300);
            w = Math.Min(w, 460) + 20;
            int h = more.Length > 0 ? 60 : 45;
            int adviceH = advice.Length > 0 ? (int)Math.Ceiling(g.MeasureString(advice, Ui.Label, w - 28, Ui.Wrap).Height) + 6 : 0;
            int adviceY = h - 2;
            if (adviceH > 0) h += adviceH + 6;
            int x = mouse.X + 16, y = mouse.Y + 18;
            if (x + w > ClientSize.Width - 6) x = Math.Max(4, mouse.X - w - 10);
            if (y + h > ClientSize.Height - 8) y = Math.Max(4, mouse.Y - h - 10);
            Ui.Shadow(g, x, y, w, h, 4, 2, 4, 38);
            Ui.Box(g, x, y, w, h, 4, Ui.White, Ui.GrayBorder);
            Ui.Text(g, n.Name, Ui.LabelBold, Ui.Graphite, x + 9, y + 6, w - 16, 16, Ui.Left);
            bool faded;
            Ui.Box(g, x + 10, y + 27, 9, 9, 1, Fill(n, out faded), Color.Empty);
            Ui.Text(g, size, Ui.LabelBold, Ui.Graphite, x + 24, y + 23, sizeW + 4, 16, Ui.Left);
            Ui.Text(g, share, Ui.Label, Ui.Gray, x + 30 + sizeW, y + 23, w - 38 - sizeW, 16, Ui.Left);
            if (more.Length > 0) Ui.Text(g, more, Ui.Caption, Ui.Gray, x + 9, y + 41, w - 16, 14, Ui.Left);
            if (adviceH > 0)
            {
                // the orange bar marks what needs attention
                g.FillRectangle(Ui.B(Ui.Orange), x + 9, y + adviceY + 3, 3, adviceH - 6);
                g.DrawString(advice, Ui.Label, Ui.B(Ui.Graphite), new RectangleF(x + 17, y + adviceY, w - 28, adviceH + 2), Ui.Wrap);
            }
        }

        void Render(float dpiX, float dpiY)
        {
            int W = ClientSize.Width, H = ClientSize.Height;
            if (W < 8 || H < 8) return;
            if (bmp == null || bmp.Width != W || bmp.Height != H)
            {
                if (bmp != null) bmp.Dispose();
                bmp = new Bitmap(W, H, PixelFormat.Format32bppPArgb);
                bmp.SetResolution(dpiX, dpiY);      // same dpi as the screen or the blit below gets stretched
            }
            hits.Clear();
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Ui.White);
                g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                if (view != null && view.Size > 0 && view.Kids != null && view.Kids.Length > 0)
                    Lay(g, view, 3, 3, W - 6, H - 6, 0);
                else
                    Ui.Text(g, view == null ? "" : "Nothing in here takes up any space", Ui.Base, Ui.Gray, 0, 0, W, H, Ui.Mid);
            }
            stale = false;
        }

        // squarified treemap - fill rows along the short side and stop adding to a row once the blocks start getting skinnier
        void Lay(Graphics g, Node dir, double x, double y, double w, double h, int depth)
        {
            Node[] kids = dir.Kids;
            double remain = dir.Size;
            int n = kids.Length;
            int i = 0;
            while (i < n && w >= 1 && h >= 1 && remain > 0)
            {
                double scale = w * h / remain;
                double side = Math.Min(w, h);
                double big = kids[i].Size * scale;
                double sum = 0, best = double.MaxValue;
                int j = i;
                while (j < n)
                {
                    double a = kids[j].Size * scale;
                    if (a <= 0) break;
                    double s2 = sum + a;
                    double worst = Math.Max(side * side * big / (s2 * s2), s2 * s2 / (side * side * a));
                    if (j > i && worst > best) break;
                    best = worst;
                    sum = s2;
                    j++;
                }
                if (j == i) break;      // only empty things left
                double thick = sum / side;
                bool wide = w >= h;
                double pos = 0;
                for (int k = i; k < j; k++)
                {
                    double len = kids[k].Size * scale / thick;
                    double rx = wide ? x : x + pos, ry = wide ? y + pos : y;
                    double rw = wide ? thick : len, rh = wide ? len : thick;
                    pos += len;
                    int x0 = (int)Math.Round(rx), y0 = (int)Math.Round(ry);
                    int x1 = (int)Math.Round(rx + rw), y1 = (int)Math.Round(ry + rh);
                    if (x1 > x0 && y1 > y0) Block(g, kids[k], x0, y0, x1 - x0, y1 - y0, depth + 1);
                    remain -= kids[k].Size;
                }
                if (wide) { x += thick; w -= thick; } else { y += thick; h -= thick; }
                i = j;
            }
        }

        void Block(Graphics g, Node n, int x, int y, int w, int h, int depth)
        {
            // neighbours are kept apart by a gap, not by borders drawn round them
            int gx = w >= 8 ? 1 : 0, gy = h >= 8 ? 1 : 0;
            int bx = x + gx, by = y + gy, bw = w - 2 * gx, bh = h - 2 * gy;
            // below 5 px a gap would eat a third of the block and the rounding makes it flicker in and out, so tiny ones touch
            if (gx == 0 && bw >= 5) bw -= 1;
            if (gy == 0 && bh >= 5) bh -= 1;
            if (bw <= 0 || bh <= 0) return;

            bool folder = Tree.HasKids(n) && n.Size > 0;
            if (!folder || bw < 12 || bh < 12)
            {
                bool faded;
                Color fill = Fill(n, out faded);
                g.FillRectangle(Ui.B(fill), bx, by, bw, bh);
                Note(n, bx, by, bw, bh);
                if (bw >= 46 && bh >= 18 && !faded)
                {
                    Color ink = Ui.InkOn(fill);
                    Ui.Text(g, n.Name, Ui.Label, ink, bx + 4, by + 1, bw - 7, 16, Ui.Left);
                    if (bh >= 34) Ui.Text(g, Fmt.Size(n.Size), Ui.Label, ink, bx + 4, by + 15, bw - 7, 16, Ui.Left);
                }
                return;
            }

            // a small folder holding nothing but one other folder would just be a frame inside a frame,
            // and a long chain of those shrinks the real content to a dot. skip straight to what is inside
            bool roomForHead = bw >= 70 && bh >= 46;
            if (!roomForHead && n.Kids.Length == 1 && Tree.HasKids(n.Kids[0]))
            {
                Block(g, n.Kids[0], x, y, w, h, depth);
                return;
            }

            // a folder is a white box with a hairline round it and its kids laid out inside. it gets no fill of its
            // own - colour in the map is for the files. top level folders get the firmer line, round corners and a bold name
            bool outer = depth == 1 && roomForHead;
            bool lined = bw >= 24 && bh >= 24;
            if (outer) Ui.Box(g, bx, by, bw, bh, 6f, Ui.White, Ui.GrayBorder);
            else if (lined) g.DrawRectangle(Ui.P(Ui.GrayLine), bx, by, bw - 1, bh - 1);
            Note(n, bx, by, bw, bh);
            int head = 0, pad = 1;
            if (roomForHead)
            {
                head = 19;
                pad = 3;
                Font face = outer ? Ui.LabelBold : Ui.Label;
                if (bw >= 150)
                {
                    Ui.Text(g, Fmt.Size(n.Size), Ui.Label, Ui.Gray, bx + bw - 80, by + 1, 74, 18, Ui.Right);
                    Ui.Text(g, n.Name, face, Ui.Graphite, bx + 6, by + 1, bw - 92, 18, Ui.Left);
                }
                else Ui.Text(g, n.Name, face, Ui.Graphite, bx + 6, by + 1, bw - 11, 18, Ui.Left);
            }
            else if (lined) pad = 2;
            int topPad = head > 0 ? head : pad;
            Lay(g, n, bx + pad, by + topPad, bw - 2 * pad, bh - topPad - pad, depth);
        }

        void Note(Node n, int x, int y, int w, int h)
        {
            Hit hit;
            hit.N = n;
            hit.X = x;
            hit.Y = y;
            hit.W = w;
            hit.H = h;
            hits.Add(hit);
        }

        Color Fill(Node n, out bool faded)
        {
            // a folder too small to open up takes the colour of the biggest file inside it
            Node c = n;
            while (c.Kind == Node.KDir && c.Kids != null && c.Kids.Length > 0) c = c.Kids[0];
            int s = -1;
            if (c.Kind != Node.KDir && slot != null && c.Type >= 0 && c.Type < slot.Length) s = slot[c.Type];
            // the guide's way of picking one out: it stays at full strength and the others dim.
            // everything e;se is drawn at that strength all the time, so it just stays as it is
            int on = Shown;
            faded = on >= 0 && s != on;
            if (s < 0) return Ui.Other;
            return faded ? Ui.TypeDim[s] : Ui.TypeFill[s];
        }

        void PaintBusy(Graphics g)
        {
            g.Clear(Ui.White);
            int W = ClientSize.Width, H = ClientSize.Height;
            int bw = Math.Max(120, Math.Min(480, W - 80));
            int x = (W - bw) / 2, y = Math.Max(20, H / 2 - 70);
            Ui.Text(g, BusyTitle, Ui.Head, Ui.Graphite, 20, y, W - 40, 28, Ui.Mid);
            if (BusyShare >= 0) Ui.Bar(g, x, y + 42, bw, 8, BusyShare, Ui.Yellow, Ui.GrayLight);
            Ui.Text(g, BusyLine1, Ui.Base, Ui.Graphite, 20, y + 60, W - 40, 20, Ui.Mid);
            Ui.Text(g, BusyLine2, Ui.Label, Ui.Gray, 20, y + 82, W - 40, 18, Ui.Mid);
            Ui.Text(g, "Esc or the Stop button ends the scan early and shows what it has so far", Ui.Label, Ui.Gray, 20, y + 112, W - 40, 18, Ui.Mid);
        }

        // each drive is one of the guide's cards: white, hairline edge, orange bar on top, a faint shadow at rest.
        // it can be clicked, so it comes up 2px and the shadow deepens while the mouse is on it
        void PaintDrives(Graphics g)
        {
            g.Clear(Ui.White);
            int W = ClientSize.Width;
            Ui.Text(g, "Pick a drive to scan", Ui.Head, Ui.Graphite, 24, 16, W - 48, 24, Ui.Left);
            Ui.Text(g, "Or type a folder path at the top and press Scan", Ui.Label, Ui.Gray, 24, 40, W - 48, 18, Ui.Left);
            int y = 74;
            int rowW = Math.Max(200, Math.Min(640, W - 48));
            for (int i = 0; i < drives.Count; i++)
            {
                DriveRow d = drives[i];
                d.Box = new Rectangle(24, y, rowW, 70);
                float t = i < rise.Length ? Ui.Ease(rise[i]) : 0f;
                float cy = y - 2f * t;
                Ui.Shadow(g, 24, cy, rowW, 70, 6, 1 + 5 * t, 2 + (int)Math.Round(8 * t), 14 + (int)(42 * t));
                Ui.Card(g, 24, cy, rowW, 70);
                double used = d.Total > 0 ? 1.0 - (double)d.Free / d.Total : 0;
                // same two steps as the free tile up top: the words turn red under 20 percent free, the bar as well under 10
                bool low = d.Total > 0 && (double)d.Free / d.Total < 0.20;
                bool tight = d.Total > 0 && (double)d.Free / d.Total < 0.10;
                string name = string.IsNullOrEmpty(d.Label) || d.Label == d.Root ? d.Root : d.Root + "   " + d.Label;
                Ui.Text(g, name, Ui.Head, Ui.Graphite, 38, cy + 10, rowW - 28, 22, Ui.Left);
                Ui.Bar(g, 38, (int)Math.Round(cy) + 37, rowW - 28, 8, used, tight ? Ui.Red : Ui.Yellow, Ui.GrayLight);
                string line = Fmt.Size(d.Free) + " free of " + Fmt.Size(d.Total);
                // bad is red and always comes with a mark in front, so it never rests on the colour alone
                if (low) line = "\u25BC Low on space   " + line;
                Ui.Text(g, line, Ui.Label, low ? Ui.RedText : Ui.Gray, 38, cy + 48, rowW - 28, 16, Ui.Left);
                y += 84;
            }
            if (drives.Count == 0) Ui.Text(g, "No fixed drives found", Ui.Base, Ui.Gray, 24, y, W - 48, 20, Ui.Left);
        }

        // ---------- mouse

        int HitAt(int x, int y)
        {
            for (int i = hits.Count - 1; i >= 0; i--)
            {
                Hit h = hits[i];
                if (x >= h.X && y >= h.Y && x < h.X + h.W && y < h.Y + h.H) return i;
            }
            return -1;
        }

        Node NodeAt(Point p)
        {
            int i = HitAt(p.X, p.Y);
            return i < 0 ? null : hits[i].N;
        }

        int DriveAt(Point p)
        {
            for (int i = 0; i < drives.Count; i++) if (drives[i].Box.Contains(p)) return i;
            return -1;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (mode == Idle)
            {
                int d = DriveAt(e.Location);
                if (d != hotDrive) { hotDrive = d; Cursor = d >= 0 ? Cursors.Hand : Cursors.Default; anim.Start(); }
                return;
            }
            if (mode != Map) return;
            mouse = e.Location;
            Node n = NodeAt(e.Location);
            if (n != hover)
            {
                hover = n;
                if (Hovered != null) Hovered(n);
            }
            Invalidate();
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            if (hotDrive != -1) { hotDrive = -1; anim.Start(); }
            if (hover != null)
            {
                hover = null;
                if (Hovered != null) Hovered(null);
                Invalidate();
            }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            Focus();
            if (mode == Idle)
            {
                int d = DriveAt(e.Location);
                if (d >= 0 && e.Button == MouseButtons.Left && DriveWanted != null) DriveWanted(drives[d].Root);
                return;
            }
            if (mode != Map) return;
            Node n = NodeAt(e.Location);
            if (n == null) return;
            pick = n;
            Invalidate();
            if (Picked != null) Picked(n);
            if (e.Button == MouseButtons.Right && MenuWanted != null) MenuWanted(n, PointToScreen(e.Location));
        }

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            base.OnMouseDoubleClick(e);
            if (mode != Map || e.Button != MouseButtons.Left || view == null) return;
            Node n = NodeAt(e.Location);
            if (n == null) return;
            // on a folder header go straight there. on a file go one level down towards it
            Node target = n;
            if (!Tree.HasKids(n))
            {
                target = null;
                for (Node p = n; p != null && p != view; p = p.Parent) if (p.Parent == view) target = p;
            }
            if (target != null && Tree.HasKids(target) && target != view && ZoomWanted != null) ZoomWanted(target);
        }
    }

    // the key under the map. point at a file type and the rest dim for as long as the mouse stays there,
    // click to keep it that way. that is the guide's rule for picking one mark out of a chart
    internal sealed class Legend : Control
    {
        string[] names = new string[0];
        long[] bytes = new long[0];
        long other;
        int emph = -1, hot = -1;
        readonly List<Rectangle> boxes = new List<Rectangle>();

        public event SlotEvent EmphasisWanted;
        public event SlotEvent PreviewWanted;

        public Legend()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.ResizeRedraw, true);
            BackColor = Ui.White;
        }

        public void Set(string[] typeNames, long[] typeBytes, long otherBytes)
        {
            names = typeNames ?? new string[0];
            bytes = typeBytes ?? new long[0];
            other = otherBytes;
            emph = -1;
            hot = -1;
            Invalidate();
        }

        public void SetEmphasis(int s) { emph = s; Invalidate(); }

        // same types in the same colours, new sizes. the highlighted type stays highlighted
        public void Sizes(long[] typeBytes, long otherBytes)
        {
            bytes = typeBytes ?? new long[0];
            other = otherBytes;
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(Ui.White);
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            g.DrawLine(Ui.P(Ui.GrayLine), 0, 0, ClientSize.Width, 0);
            boxes.Clear();
            if (names.Length == 0) return;
            int shown = hot >= 0 ? hot : emph;
            int h = 24, x = 10, y = 3;
            Ui.Text(g, "File types", Ui.Label, Ui.Gray, x, y, 64, h, Ui.Left);
            x += 68;
            // each coloured type, then everything else. wraps onto a second row
            int count = names.Length + 1;
            for (int i = 0; i < count; i++)
            {
                long size = i < names.Length ? bytes[i] : other;
                if (i >= names.Length && size <= 0) { boxes.Add(Rectangle.Empty); continue; }
                string label = i < names.Length ? names[i] : "everything else";
                string amount = Fmt.Size(size);
                Font face = i == emph ? Ui.LabelBold : Ui.Label;
                int lw = Ui.Width(g, label, face), sw = Ui.Width(g, amount, Ui.Label);
                int w = 16 + lw + 4 + sw + 14;
                if (x + w > ClientSize.Width && x > 78) { x = 78; y += h; }
                if (y + h > ClientSize.Height + 2) break;
                bool dimmed = shown >= 0 && i != shown && i < names.Length;
                Color c = i == names.Length ? Ui.Other : (dimmed ? Ui.TypeDim[i] : Ui.TypeFill[i]);
                Ui.Box(g, x, y + h / 2 - 5, 10, 10, 1, c, Color.Empty);
                Ui.Text(g, label, face, dimmed ? Ui.Gray : Ui.Graphite, x + 15, y, lw + 4, h, Ui.Left);
                Ui.Text(g, amount, Ui.Label, Ui.Gray, x + 16 + lw + 4, y, sw + 4, h, Ui.Left);
                boxes.Add(i < names.Length ? new Rectangle(x - 4, y + 2, w - 6, h - 4) : Rectangle.Empty);
                x += w;
            }
        }

        int KeyAt(Point p)
        {
            for (int i = 0; i < boxes.Count; i++) if (boxes[i].Contains(p)) return i;
            return -1;
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            int i = KeyAt(e.Location);
            if (i >= 0 && EmphasisWanted != null) EmphasisWanted(i == emph ? -1 : i);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            Aim(KeyAt(e.Location));
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            Aim(-1);
        }

        void Aim(int i)
        {
            Cursor = i >= 0 ? Cursors.Hand : Cursors.Default;
            if (i == hot) return;
            hot = i;
            Invalidate();
            if (PreviewWanted != null) PreviewWanted(i);
        }
    }

    // header strip to the guide - what got scanned on the left, number tiles on the right and
    // the 3px orange bar underneath
    internal sealed class Summary : Control
    {
        const int TileW = 140, TileH = 60, Gap = 10;
        public string Head = "", Line1 = "", Line2 = "";
        public bool Ready, HasVolume;
        public long Bytes, FileCount, DirCount, Unread, Hidden, VolTotal, VolFree;
        public string VolName = "";
        Rectangle linkBox;
        bool overLink;
        float lift;
        readonly System.Windows.Forms.Timer anim = new System.Windows.Forms.Timer();

        public event EventHandler LinkClicked;

        public Summary()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.ResizeRedraw, true);
            BackColor = Ui.White;
            anim.Interval = 15;
            anim.Tick += delegate
            {
                float want = overLink ? 1f : 0f;
                lift = lift < want ? Math.Min(want, lift + 15f / 180f) : Math.Max(want, lift - 15f / 180f);
                if (lift == want) anim.Stop();
                Invalidate();
            };
        }

        public void Clear()
        {
            Head = Line1 = Line2 = "";
            Ready = false;
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(Ui.White);
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            int W = ClientSize.Width, H = ClientSize.Height;
            // orange is the accent, and this is the bar the guide puts under a header
            g.FillRectangle(Ui.B(Ui.Orange), 0, H - 3, W, 3);
            linkBox = Rectangle.Empty;
            int textW = W - 28;
            if (Ready)
            {
                // tiles in reading order. when the window is narrow the least important ones go first
                List<int> tiles = new List<int>();
                tiles.Add(0);
                tiles.Add(1);
                tiles.Add(2);
                if (Hidden > 0) tiles.Add(3);
                if (Unread > 0) tiles.Add(4);
                if (HasVolume && VolTotal > 0) tiles.Add(5);
                int[] dropOrder = { 2, 1, 3 };
                for (int d = 0; d < dropOrder.Length && tiles.Count * (TileW + Gap) > W - 300; d++) tiles.Remove(dropOrder[d]);
                int x = W - 12 - tiles.Count * (TileW + Gap) + Gap;
                textW = x - 28;
                for (int i = 0; i < tiles.Count; i++, x += TileW + Gap) Tile(g, tiles[i], x, 10);
            }
            Ui.Text(g, Head, Ui.Head, Ui.Graphite, 14, 10, textW, 22, Ui.Left);
            Ui.Text(g, Line1, Ui.Label, Ui.Gray, 14, 34, textW, 16, Ui.Left);
            Ui.Text(g, Line2, Ui.Label, Ui.Gray, 14, 51, textW, 16, Ui.Left);
        }

        // each tile has its own colour and keeps it whichever of the others are showing. with all six up they
        // run in the guide's order: yellow, blue, green, aqua, orange, purple
        void Tile(Graphics g, int kind, int x, int y)
        {
            string value = "", label = "";
            Color labelInk = Ui.Gray;
            Font labelFace = Ui.Label;
            switch (kind)
            {
                case 0: value = Fmt.Size(Bytes); label = "scanned"; break;
                case 1: value = Fmt.Count(FileCount); label = FileCount == 1 ? "file" : "files"; break;
                case 2:
                    value = Fmt.Count(DirCount);
                    label = DirCount == 1 ? "folder" : "folders";
                    labelInk = Ui.Graphite;     // green-light is too dark for the grey, the guide says graphite only
                    break;
                case 3: value = Fmt.Size(Hidden); label = "in use but not visible"; break;
                case 4:
                    value = Fmt.Count(Unread);
                    label = "could not be read, list";
                    linkBox = new Rectangle(x, y, TileW, TileH);
                    break;
                default:
                    double free = (double)VolFree / VolTotal;
                    value = Fmt.Size(VolFree);
                    label = "free of " + Fmt.Size(VolTotal);
                    // there is no red tile. bad is red text with a mark in front so it never rests on the colour,
                    // and it goes bold when things get tight
                    if (free < 0.20) { label = "\u25BC low, " + label; labelInk = Ui.RedText; }
                    if (free < 0.10) labelFace = Ui.LabelBold;
                    break;
            }
            float fy = y;
            if (kind == 4 && lift > 0)
            {
                // anything you can click rises 2px and takes a shadow while the mouse is on it
                float t = Ui.Ease(lift);
                fy = y - 2f * t;
                Ui.Shadow(g, x, fy, TileW, TileH, 6, 1 + 3 * t, 5, (int)(56 * t));
            }
            Ui.Tile(g, x, fy, TileW, TileH, Ui.TileLight[kind], Ui.TileBar[kind]);
            Ui.Text(g, value, Ui.Numeral, Ui.Graphite, x + 8, fy + 8, TileW - 12, 28, Ui.Left);
            Ui.Text(g, label, labelFace, labelInk, x + 10, fy + 38, TileW - 14, 16, Ui.Left);
            if (kind == 4)
            {
                int ly = (int)Math.Round(fy) + 54;
                g.DrawLine(Ui.P(Ui.Orange), x + 11, ly, x + 9 + Math.Min(TileW - 22, Ui.Width(g, label, labelFace)), ly);
            }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (linkBox.Contains(e.Location) && LinkClicked != null) LinkClicked(this, EventArgs.Empty);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            Over(linkBox.Contains(e.Location));
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            Over(false);
        }

        void Over(bool inside)
        {
            Cursor = inside ? Cursors.Hand : Cursors.Default;
            if (inside == overLink) return;
            overLink = inside;
            anim.Start();
        }
    }

    // strip that shows when the process has no admin rights, so nobody takes a partial scan for the full picture.
    // orange because that is the guide's colour for what needs attention: its light background and a bar on top
    internal sealed class Notice : Control
    {
        public int Reserve;     // room kept clear on the right for the restart button

        public Notice()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.ResizeRedraw, true);
            BackColor = Ui.OrangeLight;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(Ui.OrangeLight);
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            int W = ClientSize.Width, H = ClientSize.Height;
            g.FillRectangle(Ui.B(Ui.Orange), 0, 0, W, 3);
            // warning triangle. the colour never carries it alone, the words sit right next to it
            int x = 12, y = (H + 3) / 2 - 8;
            SmoothingMode sm = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.FillPolygon(Ui.B(Ui.Orange), new Point[] { new Point(x + 8, y), new Point(x + 17, y + 16), new Point(x - 1, y + 16) });
            g.SmoothingMode = sm;
            g.FillRectangle(Ui.B(Ui.White), x + 7, y + 6, 2, 5);
            g.FillRectangle(Ui.B(Ui.White), x + 7, y + 12, 2, 2);
            Ui.Text(g, Text, Ui.Bold, Ui.OrangeText, 38, 3, W - 46 - Reserve, H - 3, Ui.Left);
        }
    }

    // a button to the guide: anything clickable rises 2px and takes a shadow under the mouse, over 0.18s.
    // sti;l a real winforms button underneath, so tab, enter and space work as they always did.
    // the control is a little bigger than the face it draws so the rise and the shadow have somewhere to go
    internal sealed class LiftButton : Button
    {
        const int PadX = 3, PadTop = 4, PadBottom = 6;
        public int Kind;        // 0 plain, 1 the main action, 2 the main action sat on a tinted strip
        bool inside, held;
        float lift;
        readonly System.Windows.Forms.Timer anim = new System.Windows.Forms.Timer();

        public LiftButton()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.ResizeRedraw, true);
            anim.Interval = 15;
            anim.Tick += delegate
            {
                float want = inside && !held && Enabled ? 1f : 0f;
                lift = lift < want ? Math.Min(want, lift + 15f / 180f) : Math.Max(want, lift - 15f / 180f);
                if (lift == want) anim.Stop();
                Invalidate();
            };
        }

        // says where the face goes. the control takes its margins on top of that
        public void Place(int x, int y, int w, int h)
        {
            SetBounds(x - PadX, y - PadTop, w + 2 * PadX, h + PadTop + PadBottom);
        }

        protected override void OnMouseEnter(EventArgs e) { base.OnMouseEnter(e); inside = true; anim.Start(); }
        protected override void OnMouseLeave(EventArgs e) { base.OnMouseLeave(e); inside = false; anim.Start(); }
        protected override void OnMouseDown(MouseEventArgs mevent) { base.OnMouseDown(mevent); held = true; anim.Start(); Invalidate(); }
        protected override void OnMouseUp(MouseEventArgs mevent) { base.OnMouseUp(mevent); held = false; anim.Start(); Invalidate(); }
        protected override void OnEnabledChanged(EventArgs e) { base.OnEnabledChanged(e); anim.Start(); Invalidate(); }
        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
        protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }

        protected override void OnPaint(PaintEventArgs pevent)
        {
            Graphics g = pevent.Graphics;
            g.Clear(Parent != null ? Parent.BackColor : Ui.White);
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            float t = Ui.Ease(lift);
            float fx = PadX, fw = Width - 2 * PadX, fh = Height - PadTop - PadBottom;
            float fy = PadTop - 2f * t;
            if (t > 0) Ui.Shadow(g, fx, fy, fw, fh, 4, 1 + 2 * t, 3, (int)(60 * t));

            // plain ones are white with a gray-border edge. the main action wears the accent, orange with its own
            // light background and text step, or white inside when it already sits on orange-light
            bool main = Kind != 0;
            Color fill = !main ? (held ? Ui.GrayLight : Ui.White)
                : (Kind == 1) != held ? Ui.OrangeLight : Ui.White;
            Color edge = main ? Ui.Orange : Ui.GrayBorder;
            Color ink = main ? Ui.OrangeText : Ui.Graphite;
            if (!Enabled) { fill = Ui.White; edge = Ui.GrayLine; ink = Ui.GrayBorder; }
            Ui.Box(g, fx, fy, fw, fh, 4, fill, edge);
            Ui.Text(g, Text, main ? Ui.Bold : Ui.Base, ink, fx + 2, fy, fw - 4, fh, Ui.Mid);
            // keyboard focus gets the guide's blue outline
            if (Focused && ShowFocusCues)
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (GraphicsPath ring = Ui.Round(fx + 2.5f, fy + 2.5f, fw - 5, fh - 5, 2, 2, 2, 2)) g.DrawPath(Ui.P(Ui.Blue), ring);
                g.SmoothingMode = SmoothingMode.None;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) anim.Dispose();
            base.Dispose(disposing);
        }
    }

    internal sealed class Options
    {
        public string Path = "";
        public string Script = "";
        public bool Backup;
        public int PerFolder = 5000;
        public int Threads;
        public bool Classic;
        public bool ReadOnly;
    }

    internal sealed class MainForm : Form
    {
        readonly Options opt;
        readonly ComboBox pathBox = new ComboBox();
        readonly LiftButton browseBtn = new LiftButton(), scanBtn = new LiftButton(), upBtn = new LiftButton(),
            openBtn = new LiftButton(), adminBtn = new LiftButton();
        readonly Label pathLbl = new Label(), crumb = new Label(), status = new Label();
        readonly Summary summary = new Summary();
        readonly Notice notice = new Notice();
        readonly TreeGrid tree = new TreeGrid();
        readonly MapView map = new MapView();
        readonly Legend legend = new Legend();
        readonly Panel grip = new Panel();
        readonly ContextMenuStrip menu = new ContextMenuStrip();
        readonly System.Windows.Forms.Timer tick = new System.Windows.Forms.Timer();

        Scanner scan, sub;                  // sub is a rescan of one folder inside the tree
        Node root, view, subNode, leftover;
        readonly List<Node> redo = new List<Node>();        // folders waiting to be scanned again
        readonly List<string> unread = new List<string>();
        int unreadCount, removed;
        bool shown;
        TypeTable types;
        sbyte[] typeSlot = new sbyte[0];    // colour slot per file type, -1 for everything else
        int[] slotType = new int[0];        // and back the other way for the key under the map
        string after = "";                  // what the status line says once a rescan has settled
        Stopwatch clock = new Stopwatch();
        int split = -1;
        bool dragging;
        readonly bool elevated, backupOn, notAdmin, canRestart;
        long volTotal, volFree;
        bool volKnown, wholeVolume;
        const string Hint = "Double click a block to zoom in. Backspace goes up a level. Right click for more.";

        public MainForm(Options o)
        {
            opt = o;
            elevated = Native.IsElevated();
            backupOn = o.Backup && Native.EnableBackupPrivilege();

            notAdmin = Native.IsWindows && !elevated;
            canRestart = notAdmin && !string.IsNullOrEmpty(o.Script);
            Text = "OOutSize" + (notAdmin ? "  (not running as administrator)" : "") + (o.ReadOnly ? "  (read only)" : "");
            AutoScaleMode = AutoScaleMode.None;
            Font = Ui.Base;
            BackColor = Ui.White;
            StartPosition = FormStartPosition.CenterScreen;
            Rectangle area = Screen.PrimaryScreen.WorkingArea;
            ClientSize = new Size(Math.Max(760, Math.Min(1360, area.Width - 60)), Math.Max(480, Math.Min(860, area.Height - 80)));
            MinimumSize = new Size(760, 480);
            KeyPreview = true;

            pathLbl.Text = "Path";
            pathLbl.Font = Ui.Label;
            pathLbl.ForeColor = Ui.Gray;
            menu.Font = Ui.Base;
            pathBox.DropDownStyle = ComboBoxStyle.DropDown;
            try
            {
                foreach (DriveInfo d in DriveInfo.GetDrives())
                    if (d.DriveType == DriveType.Fixed || d.DriveType == DriveType.Network || d.DriveType == DriveType.Removable)
                        pathBox.Items.Add(d.Name);
            }
            catch { }
            pathBox.KeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.KeyCode != Keys.Enter) return;
                e.Handled = true;
                e.SuppressKeyPress = true;
                StartScan(pathBox.Text);
            };
            pathBox.SelectionChangeCommitted += delegate
            {
                // only when a drive was clicked in the open list. a closed combo also changes its pick on
                // the mouse wheel and arrow keys and nobody wants a scan kicked off by a stray scroll
                object picked = pathBox.SelectedItem;
                if (picked != null && pathBox.DroppedDown) BeginInvoke(new TextEvent(StartScan), picked.ToString());
            };

            Setup(browseBtn, "Browse...", false, delegate { Browse(); });
            Setup(scanBtn, "Scan", true, delegate { if (Scanning) StopScan(); else StartScan(pathBox.Text); });
            Setup(upBtn, "Up", false, delegate { GoUp(); });
            Setup(openBtn, "Open in Explorer", false, delegate { if (view != null) Reveal(view); });
            Setup(adminBtn, "Restart as admin", true, delegate { Elevate(); });
            adminBtn.Kind = 2;
            adminBtn.Visible = canRestart;
            notice.Text = "Not running as administrator. Folders this account cannot read are left out, so the totals may be short.";
            notice.Reserve = canRestart ? 148 : 0;
            notice.Visible = notAdmin;
            notice.Controls.Add(adminBtn);
            upBtn.Enabled = openBtn.Enabled = false;

            crumb.AutoEllipsis = true;
            crumb.TextAlign = ContentAlignment.MiddleLeft;
            crumb.ForeColor = Ui.Graphite;
            crumb.BackColor = Ui.White;
            status.AutoEllipsis = true;
            status.TextAlign = ContentAlignment.MiddleLeft;
            status.ForeColor = Ui.Gray;
            status.Text = Hint;
            // the divider is a hairline now. orange is kept for the bar under the header and for what needs attention
            grip.Cursor = Cursors.VSplit;
            grip.BackColor = Ui.White;
            grip.Paint += delegate(object s, PaintEventArgs e) { e.Graphics.DrawLine(Ui.P(Ui.GrayBorder), 2, 0, 2, grip.Height); };
            grip.MouseDown += delegate { dragging = true; grip.Capture = true; };
            grip.MouseUp += delegate { dragging = false; grip.Capture = false; };
            grip.MouseMove += delegate(object s, MouseEventArgs e)
            {
                if (!dragging) return;
                split = PointToClient(grip.PointToScreen(e.Location)).X - 2;
                Arrange();
            };

            tree.Picked += TreePicked;
            tree.MenuWanted += ShowMenu;
            map.Picked += delegate(Node n) { tree.Select(n, false); Describe(n); };
            map.Hovered += delegate(Node n) { Describe(n ?? tree.Selected); };
            map.ZoomWanted += delegate(Node n) { tree.Select(n, false); SetView(n); };
            map.MenuWanted += ShowMenu;
            map.DriveWanted += StartScan;
            legend.EmphasisWanted += delegate(int s) { legend.SetEmphasis(s); map.SetEmphasis(s); };
            legend.PreviewWanted += delegate(int s) { map.SetPreview(s); };
            summary.LinkClicked += delegate { ShowUnreadable(); };
            tick.Interval = 150;
            tick.Tick += delegate { OnTick(); };

            Controls.AddRange(new Control[] { pathLbl, pathBox, browseBtn, scanBtn, upBtn, openBtn, notice,
                summary, tree, grip, crumb, map, legend, status });
            Arrange();
            map.ShowIdle();
        }

        // the main action wears the accent, the rest are plain. how they look and move is in LiftButton
        void Setup(LiftButton b, string text, bool main, EventHandler click)
        {
            b.Text = text;
            b.Kind = main ? 1 : 0;
            b.Click += click;
        }

        // the footer is a hairline with grey text under it, the way the guide does one
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            int y = ClientSize.Height - 26;
            e.Graphics.DrawLine(Ui.P(Ui.GrayLine), 0, y, ClientSize.Width, y);
        }

        bool Scanning { get { return sub != null || (scan != null && !scan.Done); } }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            Arrange();
            if (!string.IsNullOrEmpty(opt.Path)) StartScan(opt.Path);
            else pathBox.Focus();
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            Arrange();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (scan != null) scan.Cancel = true;
            if (sub != null) sub.Cancel = true;
            tick.Stop();
            base.OnFormClosing(e);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.Handled) return;
            if (e.KeyCode == Keys.Escape && Scanning) { StopScan(); e.Handled = true; }
            else if (e.KeyCode == Keys.F5 && !Scanning && pathBox.Text.Trim().Length > 0) { StartScan(pathBox.Text); e.Handled = true; }
            else if (e.KeyCode == Keys.Back && !pathBox.Focused && view != null) { GoUp(); e.Handled = true; }
            else if (e.KeyCode == Keys.Delete && !pathBox.Focused && !opt.ReadOnly && root != null && tree.Selected != null)
            {
                Delete(tree.Selected, e.Shift);
                e.Handled = true;
            }
        }

        // laid out by hand. dock order in winforms is a guessing game and this is th same everywhere
        void Arrange()
        {
            int W = ClientSize.Width, H = ClientSize.Height;
            if (W < 200 || H < 200) return;
            int right = W - 10;
            pathLbl.SetBounds(12, 14, 34, 16);
            int fixedW = 48 + 8 + 86 + 6 + 72 + 22 + 56 + 6 + 128 + 16;
            int boxW = Math.Max(160, Math.Min(560, right - fixedW));
            int x = 48;
            pathBox.SetBounds(x, 10, boxW, 24); x += boxW + 8;
            browseBtn.Place(x, 8, 86, 26); x += 92;
            scanBtn.Place(x, 8, 72, 26); x += 94;
            upBtn.Place(x, 8, 56, 26); x += 62;
            openBtn.Place(x, 8, 128, 26);

            int y = 42;
            if (notAdmin)
            {
                notice.SetBounds(0, y, W, 38);
                if (canRestart) adminBtn.Place(W - 142, 8, 130, 24);
                y += 38;
            }
            summary.SetBounds(0, y, W, 84);
            int top = y + 84, bottom = H - 26, bodyH = Math.Max(60, bottom - top);
            if (split < 0) split = (int)(W * 0.47);
            int lo = 260, hi = Math.Max(lo, W - 340);
            if (split < lo) split = lo;
            if (split > hi) split = hi;
            tree.SetBounds(0, top, split, bodyH);
            grip.SetBounds(split, top, 5, bodyH);
            int rx = split + 5, rw = W - rx;
            crumb.SetBounds(rx, top, rw, 26);
            map.SetBounds(rx, top + 26, rw, Math.Max(20, bodyH - 26 - 52));
            legend.SetBounds(rx, top + bodyH - 52, rw, 52);
            status.SetBounds(10, H - 24, W - 20, 20);
            Invalidate();
        }

        // ---------- scanning

        static string Tidy(string raw)
        {
            string p = (raw ?? "").Trim().Trim('"').Trim();
            if (p.Length == 0) return "";
            if (p.Length == 2 && p[1] == ':') p += "\\";
            p = Path.GetFullPath(p);
            string rootPart = Path.GetPathRoot(p) ?? "";
            if (p.Length > rootPart.Length) p = p.TrimEnd('\\', '/');
            return p;
        }

        void StartScan(string raw)
        {
            if (Scanning) return;
            string p;
            try { p = Tidy(raw); }
            catch (Exception ex) { Say("That path does not look right.\r\n\r\n" + ex.Message); return; }
            if (p.Length == 0) return;
            if (!Directory.Exists(p)) { Say("Cannot find the folder\r\n\r\n" + p); return; }

            pathBox.Text = p;
            root = view = null;
            shown = false;
            removed = 0;
            unreadCount = 0;
            unread.Clear();
            redo.Clear();
            after = "";
            leftover = null;
            tree.Slot = null;
            tree.SetRoot(null);
            legend.Set(null, null, 0);
            summary.Clear();
            crumb.Text = "";
            upBtn.Enabled = openBtn.Enabled = false;
            volKnown = VolumeSpace(p, out volTotal, out volFree);
            string rootPart = Path.GetPathRoot(p) ?? "";
            wholeVolume = string.Equals(rootPart.TrimEnd('\\', '/'), p.TrimEnd('\\', '/'), StringComparison.OrdinalIgnoreCase);

            scan = new Scanner();
            scan.RootPath = p;
            scan.PerFolder = opt.PerFolder;
            scan.Threads = opt.Threads;
            scan.Classic = opt.Classic;
            clock = Stopwatch.StartNew();
            scan.Start();
            scanBtn.Text = "Stop";
            status.Text = "Scanning " + p;
            map.ShowBusy();
            ShowProgress();
            tick.Start();
        }

        void StopScan()
        {
            redo.Clear();
            if (sub != null) sub.Cancel = true;
            else if (scan != null) scan.Cancel = true;
        }

        void OnTick()
        {
            if (sub != null)
            {
                if (!sub.Done)
                {
                    Tell("Scanning " + sub.RootPath + " again, " + Fmt.Things(Interlocked.Read(ref sub.FilesSeen), "file", "files") + " so far");
                    return;
                }
                Scanner done = sub;
                Node was = subNode;
                sub = null;
                subNode = null;
                Splice(was, done);
                Next();
                return;
            }
            if (scan == null || shown) { tick.Stop(); return; }
            if (!scan.Done) { ShowProgress(); return; }
            tick.Stop();
            Finish();
        }

        void Tell(string text)
        {
            status.Text = text.Replace("&", "&&");
        }

        void ShowProgress()
        {
            long files = Interlocked.Read(ref scan.FilesSeen), dirs = Interlocked.Read(ref scan.DirsSeen),
                bytes = Interlocked.Read(ref scan.BytesSeen);
            map.BusyTitle = "Scanning " + scan.RootPath;
            map.BusyLine1 = Fmt.Size(bytes) + " in " + Fmt.Things(files, "file", "files") + " and "
                + Fmt.Things(dirs, "folder", "folders") + " after " + Fmt.Secs(clock.Elapsed.TotalSeconds);
            map.BusyLine2 = scan.Where;
            long used = volTotal - volFree;
            map.BusyShare = wholeVolume && volKnown && used > 0 ? Math.Min(0.99, (double)bytes / used) : -1;
            map.Invalidate();
        }

        void Finish()
        {
            Scanner s = scan;
            shown = true;
            scanBtn.Text = "Scan";
            if (s.Error != null) Say("The scan hit a problem.\r\n\r\n" + s.Error);
            root = s.Root;
            if (root == null) { map.ShowIdle(); status.Text = Hint; return; }
            types = s.Types;
            lock (s.Unreadable) unread.AddRange(s.Unreadable);
            unreadCount = s.UnreadableCount;

            // the file types using the most bytes get a colour for this scan and keep it while you zoom
            // around, delete things or scan a folder again
            string[] all = types.Names();
            long[] sums = new long[all.Length];
            Scanner.Tally(root, sums);
            typeSlot = new sbyte[all.Length];
            for (int i = 0; i < all.Length; i++) typeSlot[i] = -1;
            List<string> names = new List<string>();
            List<int> picked = new List<int>();
            for (int k = 0; k < Ui.TypeFill.Length; k++)
            {
                int best = -1;
                for (int i = 0; i < all.Length; i++)
                    if (typeSlot[i] < 0 && sums[i] > 0 && (best < 0 || sums[i] > sums[best])) best = i;
                if (best < 0) break;
                typeSlot[best] = (sbyte)k;
                names.Add(all[best].Length == 0 ? "no extension" : all[best]);
                picked.Add(best);
            }
            slotType = picked.ToArray();
            tree.Slot = typeSlot;
            legend.Set(names.ToArray(), new long[names.Count], 0);
            tree.SetRoot(root);
            map.ShowMap(root, typeSlot);
            SetView(root);
            summary.Head = s.RootPath;
            summary.VolName = Path.GetPathRoot(s.RootPath) ?? "";
            Totals();
            status.Text = Hint;
            tree.Focus();
        }

        // every number that depends on what is in the tree right now. runs after a scan and again
        // whenever something was deleted or a folder was scanned again
        void Totals()
        {
            if (root == null || scan == null || types == null) return;
            Scanner s = scan;
            string[] all = types.Names();
            if (all.Length > typeSlot.Length)
            {
                // a rescan met file types the first pass never saw. they go in with everything e;se
                int had = typeSlot.Length;
                Array.Resize(ref typeSlot, all.Length);
                for (int i = had; i < all.Length; i++) typeSlot[i] = -1;
                tree.Slot = typeSlot;
                map.SetSlots(typeSlot);
            }
            long[] sums = new long[all.Length];
            Scanner.Tally(root, sums);
            long[] sizes = new long[slotType.Length];
            long coloured = 0;
            for (int i = 0; i < slotType.Length; i++) { sizes[i] = sums[slotType[i]]; coloured += sizes[i]; }
            legend.Sizes(sizes, Math.Max(0, root.Size - coloured));

            volKnown = VolumeSpace(s.RootPath, out volTotal, out volFree);
            long used = volTotal - volFree;
            summary.Line1 = "Scanned in " + Fmt.Secs(s.Seconds) + " on " + s.Crew + " threads";
            List<string> notes = new List<string>();
            if (s.Stopped) notes.Add("Stopped early so sizes are incomplete");
            if (s.Capped) notes.Add("Huge scan, small files are grouped more tightly");
            if (removed > 0) notes.Add(Fmt.Things(removed, "item", "items") + " deleted since the scan");
            // the yellow strip already says when we are not admin, only the backup switch needs a line of its own
            if (backupOn) notes.Add("Backup privilege on");
            else if (opt.Backup) notes.Add(notAdmin ? "The backup privilege needs an admin window" : "The backup privilege could not be turned on");
            summary.Line2 = string.Join(".  ", notes.ToArray()) + (notes.Count > 0 ? "." : "");
            summary.Bytes = root.Size;
            summary.FileCount = root.Files;
            summary.DirCount = root.Dirs;
            summary.Unread = unreadCount;
            summary.Hidden = wholeVolume && volKnown && !s.Stopped && used - root.Size > Math.Max(256L << 20, used / 200) ? used - root.Size : 0;
            summary.HasVolume = volKnown;
            summary.VolTotal = volTotal;
            summary.VolFree = volFree;
            summary.Ready = true;
            summary.Invalidate();
        }

        static bool VolumeSpace(string path, out long total, out long free)
        {
            total = 0;
            free = 0;
            try
            {
                if (Native.IsWindows)
                {
                    ulong mine, all, spare;
                    if (!Native.DiskFree(path.EndsWith("\\", StringComparison.Ordinal) ? path : path + "\\", out mine, out all, out spare)) return false;
                    total = (long)all;
                    free = (long)spare;
                    return total > 0;
                }
                DriveInfo d = new DriveInfo(Path.GetPathRoot(path));
                total = d.TotalSize;
                free = d.TotalFreeSpace;
                return total > 0;
            }
            catch { return false; }
        }

        // ---------- moving around

        void TreePicked(Node n)
        {
            if (n == null) return;
            if (n.Kind == Node.KDir) { SetView(n); map.SetPick(null); }
            else { SetView(n.Parent); map.SetPick(n); }
            Describe(n);
        }

        void SetView(Node dir)
        {
            if (dir == null || dir == view) return;
            view = dir;
            map.SetView(dir);
            string text = Tree.PathOf(dir) + "     " + Fmt.Size(dir.Size);
            if (root != null && dir != root && root.Size > 0) text += ", " + Fmt.Pct((double)dir.Size / root.Size) + " of the scan";
            crumb.Text = text.Replace("&", "&&");
            upBtn.Enabled = dir.Parent != null;
            openBtn.Enabled = true;
            Describe(dir);
        }

        void GoUp()
        {
            if (view == null || view.Parent == null) return;
            Node was = view;
            SetView(view.Parent);
            tree.Select(view, false);
            map.SetPick(was);
        }

        void Describe(Node n)
        {
            if (n == null) { status.Text = Hint; return; }
            string text = Tree.PathOf(n) + "     " + Fmt.Size(n.Size);
            if (n.Kind == Node.KPile) text = Tree.PathOf(n) + Tree.Sep + "  " + n.Name + "     " + Fmt.Size(n.Size);
            if (n.Kind == Node.KDir) text += ", " + Fmt.Things(n.Files, "file", "files");
            string when = Fmt.Date(n.Stamp);
            if (when.Length > 0) text += (n.Kind == Node.KFile ? ", changed " : ", newest ") + when;
            if (n.Guard != 0 && n.Guard < Guard.Advice.Length) text += "     " + Guard.Advice[n.Guard];
            status.Text = text.Replace("&", "&&");
        }

        void ShowMenu(Node n, Point at)
        {
            if (n == null) return;
            menu.Items.Clear();
            Node folder = n.Kind == Node.KDir ? n : n.Parent;
            int shown = 0;
            for (Node p = folder; p != null && shown < 5; p = p.Parent)
            {
                if (p == view) continue;
                Node target = p;        // a for loop shares its variable between passes, the handler needs its own
                menu.Items.Add("Show " + Short(p.Name) + " in the map", null, delegate { tree.Select(target, false); SetView(target); });
                shown++;
            }
            if (shown > 0) menu.Items.Add(new ToolStripSeparator());
            Node it = n.Kind == Node.KPile ? n.Parent : n;
            menu.Items.Add("Open in Explorer", null, delegate { Reveal(it); });
            menu.Items.Add("Copy path", null, delegate { Copy(Tree.PathOf(it)); });
            if (folder != null) menu.Items.Add("Copy the listing of " + Short(folder.Name), null, delegate { Copy(Listing(folder)); });
            if (folder != null && folder != root) menu.Items.Add("Scan " + Short(folder.Name) + " again", null, delegate { Rescan(folder); });
            if (!opt.ReadOnly && n != root && n.Kind != Node.KPile)
            {
                menu.Items.Add(new ToolStripSeparator());
                if (Guard.HandsOff(n.Guard) || Guard.Keep(Tree.PathOf(n)).Length > 0)
                    menu.Items.Add("Why can this not be deleted?", null, delegate { Delete(n, false); });
                else
                {
                    // both always ask first. the dots are the windows way of saying a question follows
                    menu.Items.Add("Move to Recycle Bin...", null, delegate { Delete(n, false); });
                    menu.Items.Add("Delete permanently...", null, delegate { Delete(n, true); });
                }
            }
            menu.Show(at);
        }

        static string Short(string name)
        {
            string s = name.Length > 40 ? name.Substring(0, 37) + "..." : name;
            return s.Replace("&", "&&");
        }

        static string Listing(Node dir)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(Tree.PathOf(dir)).Append("\t").Append(Fmt.Size(dir.Size)).Append("\r\n");
            Node[] kids = dir.Kids ?? Tree.None;
            for (int i = 0; i < kids.Length; i++)
            {
                Node k = kids[i];
                sb.Append(k.Name).Append(k.Kind == Node.KDir ? "\\" : "").Append("\t").Append(Fmt.Size(k.Size)).Append("\t")
                    .Append(Fmt.Pct(dir.Size > 0 ? (double)k.Size / dir.Size : 0)).Append("\t")
                    .Append(k.Kind == Node.KFile ? "" : Fmt.Count(k.Files)).Append("\r\n");
            }
            return sb.ToString();
        }

        // ---------- deleting, and scanning one folder again

        void Delete(Node n, bool forGood)
        {
            if (n == null || root == null || opt.ReadOnly) return;
            if (Scanning) { Say("Wait for the scan to finish first."); return; }
            if (n == root) { Say("OOutSize does not delete the folder a scan started in."); return; }
            if (n.Kind == Node.KPile) { Say("That row is many small files rolled into one. Open the folder in Explorer to pick through them."); return; }
            string path = Tree.PathOf(n);
            bool isDir = n.Kind == Node.KDir;
            string what = isDir ? "folder" : "file";

            // windows owns it, or the machine falls over without it. no way round this one
            string no = Guard.HandsOff(n.Guard) && n.Guard < Guard.Advice.Length ? Guard.Advice[n.Guard] : Guard.Keep(path);
            if (no.Length > 0)
            {
                MessageBox.Show(this, "OOutSize will not delete this " + what + ".\r\n\r\n" + path + "\r\n\r\n" + no,
                    "OOutSize", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            if (!Native.IsWindows) { Say("Deleting only works on Windows."); return; }
            if (!Native.Exists(path))
            {
                Remove(n);
                Tell("Already gone: " + path);
                return;
            }
            if (!Native.Plain(path))
            {
                Say("OOutSize will not delete this " + what + ".\r\n\r\n" + path + "\r\n\r\nIts name ends in a dot or a space, uses a "
                    + "reserved name, or the path is longer than 259 characters. The Windows delete function tidies names like that "
                    + "before it acts and could end up on a different item.");
                return;
            }

            if (isDir && Longest(n) > 259)
            {
                Say("OOutSize will not delete this folder.\r\n\r\n" + path + "\r\n\r\nSomething inside it has a path longer than "
                    + "259 characters and the Windows delete function cannot handle those. Remove the deep part another way first.");
                return;
            }

            string care = n.Guard != 0 && n.Guard < Guard.Advice.Length ? Guard.Advice[n.Guard] : Guard.Care(path);
            StringBuilder q = new StringBuilder();
            if (care.Length > 0) q.Append("Careful. ").Append(care).Append("\r\n\r\n");
            if (forGood) q.Append("Permanently delete this ").Append(what).Append(isDir ? " and everything in it?" : "?");
            else q.Append("Move this ").Append(what).Append(" to the Recycle Bin?");
            q.Append("\r\n\r\n").Append(path).Append("\r\n\r\nWhen scanned: ").Append(Fmt.Size(n.Size));
            if (isDir)
            {
                q.Append(" in ").Append(Fmt.Things(n.Files, "file", "files"));
                if (n.Dirs > 0) q.Append(" and ").Append(Fmt.Things(n.Dirs, "folder", "folders"));
            }
            q.Append("\r\n\r\n");
            q.Append(forGood ? "This cannot be undone." : "The space is not freed until the Recycle Bin is emptied.");
            DialogResult answer = MessageBox.Show(this, q.ToString(), forGood ? "Delete permanently" : "Move to Recycle Bin",
                MessageBoxButtons.YesNo, forGood || care.Length > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);
            if (answer != DialogResult.Yes) return;

            long size = n.Size;
            bool aborted;
            try { Native.ShellDelete(Handle, path, !forGood, out aborted); }
            catch (Exception ex) { Say("Could not start the delete.\r\n\r\n" + ex.Message); return; }

            if (!Native.Exists(path))
            {
                removed++;
                after = (forGood ? "Deleted " : "Moved to the Recycle Bin: ") + path + "     " + Fmt.Size(size);
                // on a whole drive scan the bin is in the tree, so scan it again and the bytes show up where they went
                Node bin = forGood ? null : Bin();
                Remove(n);
                if (bin != null) Rescan(bin);
            }
            else if (isDir)
            {
                // cancelled, or some of it was in use. scan what is left so the numbers are true again
                leftover = n;
                Rescan(n);
            }
            else after = "Not deleted: " + path;
            if (!Scanning && after.Length > 0) { Tell(after); after = ""; }
            tree.Focus();
        }

        // longest full path anywhere under a row, going by what the scan saw. the files rolled into a pile
        // have no names kept, so those cannot be counted
        static int Longest(Node n)
        {
            int best = 0;
            List<Node> stack = new List<Node>();
            List<int> lengths = new List<int>();
            stack.Add(n);
            lengths.Add(Tree.PathOf(n).Length);
            while (stack.Count > 0)
            {
                int last = stack.Count - 1;
                Node d = stack[last];
                int len = lengths[last];
                stack.RemoveAt(last);
                lengths.RemoveAt(last);
                if (len > best) best = len;
                Node[] kids = d.Kids;
                if (kids == null) continue;
                for (int i = 0; i < kids.Length; i++)
                {
                    if (kids[i].Kind == Node.KPile) continue;
                    stack.Add(kids[i]);
                    lengths.Add(len + 1 + kids[i].Name.Length);
                }
            }
            return best;
        }

        Node Bin()
        {
            if (!wholeVolume || root == null || root.Kids == null) return null;
            Node[] kids = root.Kids;
            for (int i = 0; i < kids.Length; i++)
                if (kids[i].Kind == Node.KDir && string.Equals(kids[i].Name, "$Recycle.Bin", StringComparison.OrdinalIgnoreCase)) return kids[i];
            return null;
        }

        // take a row out of the tree after it was deleted on disk
        void Remove(Node n)
        {
            Node up = n.Parent;
            int at = up == null || up.Kids == null ? -1 : Array.IndexOf(up.Kids, n);
            if (at < 0) return;
            string path = Tree.PathOf(n);
            Node show = Tree.Under(view, n) ? up : view;
            Node keep = Tree.Under(tree.Selected, n) ? up : tree.Selected;
            Node[] rest = new Node[up.Kids.Length - 1];
            Array.Copy(up.Kids, 0, rest, 0, at);
            Array.Copy(up.Kids, at + 1, rest, at, rest.Length - at);
            up.Kids = rest;
            n.Parent = null;        // cut loose, so anything sti;l holding it can tell it is no longer part of the tree
            Forget(path);
            Tree.Recount(up);
            Reshow(show, keep);
        }

        // could-not-read entries under a path that is gone or about to be listed afresh
        void Forget(string path)
        {
            int had = unread.Count;
            string inside = path.EndsWith("\\", StringComparison.Ordinal) || path.EndsWith("/", StringComparison.Ordinal) ? path : path + Tree.Sep;
            for (int i = unread.Count - 1; i >= 0; i--)
            {
                string u = unread[i];
                int cut = u.IndexOf("   (error ", StringComparison.Ordinal);
                if (cut >= 0) u = u.Substring(0, cut);
                if (string.Equals(u, path, StringComparison.OrdinalIgnoreCase) || u.StartsWith(inside, StringComparison.OrdinalIgnoreCase)) unread.RemoveAt(i);
            }
            unreadCount = Math.Max(unread.Count, unreadCount - (had - unread.Count));
        }

        void Rescan(Node dir)
        {
            if (dir == null || root == null || dir.Kind != Node.KDir || dir == root) return;
            if (scan != null && !scan.Done) return;
            if (dir != subNode && !redo.Contains(dir)) redo.Add(dir);
            if (sub == null) Next();
        }

        // start on the next folder in the queue, or settle down when there is none
        void Next()
        {
            while (sub == null && redo.Count > 0)
            {
                Node dir = redo[0];
                redo.RemoveAt(0);
                if (!Tree.Under(dir, root)) continue;       // went with something deleted in the meantime
                string p = Tree.PathOf(dir);
                if (!Native.Exists(p)) { Remove(dir); continue; }
                subNode = dir;
                sub = new Scanner();
                sub.RootPath = p;
                sub.GuardRoot = scan.RootPath;
                sub.Types = types;
                sub.PerFolder = opt.PerFolder;
                sub.Threads = opt.Threads;
                sub.Classic = opt.Classic;
                sub.Start();
                scanBtn.Text = "Stop";
                Tell("Scanning " + p + " again");
                tick.Start();
            }
            if (sub != null) return;
            tick.Stop();
            scanBtn.Text = "Scan";
            leftover = null;
            if (after.Length > 0) { Tell(after); after = ""; }
        }

        // swap the fresh copy of a folder in where the old one sat
        void Splice(Node old, Scanner done)
        {
            Node fresh = done.Root;
            Node up = old == null ? null : old.Parent;
            int at = up == null || up.Kids == null ? -1 : Array.IndexOf(up.Kids, old);
            if (done.Stopped || done.Error != null || fresh == null || at < 0 || !Tree.Under(old, root))
            {
                after = "Stopped. The sizes under " + done.RootPath + " may be out of date.";
                return;
            }
            Node show = Twin(view, old, fresh), keep = Twin(tree.Selected, old, fresh);
            if (old == leftover)
            {
                bool same = fresh.Size == old.Size && fresh.Files == old.Files && fresh.Dirs == old.Dirs;
                if (!same) removed++;
                after = same ? "Nothing was deleted from " + done.RootPath
                    : "Only part of " + done.RootPath + " was deleted. " + Fmt.Size(fresh.Size) + " is left";
            }
            else if (after.Length == 0) after = "Scanned " + done.RootPath + " again     " + Fmt.Size(fresh.Size);
            fresh.Name = old.Name;
            fresh.Parent = up;
            fresh.Open = old.Open;
            up.Kids[at] = fresh;
            old.Parent = null;
            Forget(done.RootPath);
            lock (done.Unreadable) unread.AddRange(done.Unreadable);
            unreadCount += done.UnreadableCount;
            Tree.Recount(up);
            Reshow(show, keep);
        }

        // the same spot in the fresh copy of a folder, or as near to it as sti;l exists
        static Node Twin(Node n, Node old, Node fresh)
        {
            if (n == null || !Tree.Under(n, old)) return n;
            List<Node> chain = new List<Node>();
            for (Node p = n; p != old; p = p.Parent) chain.Add(p);
            Node spot = fresh;
            for (int i = chain.Count - 1; i >= 0; i--)
            {
                Node want = chain[i], found = null;
                Node[] kids = spot.Kids ?? Tree.None;
                for (int k = 0; k < kids.Length && found == null; k++)
                    if (kids[k].Kind == want.Kind && string.Equals(kids[k].Name, want.Name, StringComparison.Ordinal)) found = kids[k];
                if (found == null) break;
                spot = found;
            }
            return spot;
        }

        // the tree changed underneath - bring the tiles, the tree, the map and the key back in line with it
        void Reshow(Node show, Node keep)
        {
            Totals();
            tree.Reload(keep);
            Node target = show != null && Tree.Under(show, root) ? show : root;
            if (target.Kind != Node.KDir) target = target.Parent ?? root;
            view = null;
            SetView(target);
            Node pick = tree.Selected;
            map.SetPick(pick != null && pick.Kind != Node.KDir ? pick : null);
            Describe(pick);
        }

        void Copy(string text)
        {
            try { Clipboard.SetText(text); status.Text = "Copied to the clipboard"; }
            catch (Exception ex) { Say("Could not reach the clipboard.\r\n\r\n" + ex.Message); }
        }

        void Reveal(Node n)
        {
            try
            {
                string p = Tree.PathOf(n);
                if (n.Kind == Node.KFile) { Process.Start("explorer.exe", "/select,\"" + p + "\""); return; }
                // a quoted path ending in a backslash swallows the closing quote, so drive roots go in bare
                if (p.Length > 3) p = "\"" + p.TrimEnd('\\') + "\"";
                Process.Start("explorer.exe", p);
            }
            catch (Exception ex) { Say("Could not open Explorer.\r\n\r\n" + ex.Message); }
        }

        void Browse()
        {
            using (FolderBrowserDialog d = new FolderBrowserDialog())
            {
                d.Description = "Folder to scan";
                d.ShowNewFolderButton = false;
                if (d.ShowDialog(this) == DialogResult.OK) StartScan(d.SelectedPath);
            }
        }

        void Elevate()
        {
            try
            {
                string args = "-NoProfile -ExecutionPolicy Bypass -File " + Quote(opt.Script);
                string p = scan != null ? scan.RootPath : pathBox.Text.Trim();
                if (p.Length > 0) args += " -Path " + Quote(p);
                if (opt.Backup) args += " -BackupPrivilege";
                if (opt.PerFolder != 5000) args += " -MaxFilesPerFolder " + opt.PerFolder;
                if (opt.Threads > 0) args += " -Threads " + opt.Threads;
                if (opt.Classic) args += " -ClassicScan";
                if (opt.ReadOnly) args += " -ReadOnly";
                ProcessStartInfo si = new ProcessStartInfo("powershell.exe", args);
                si.UseShellExecute = true;
                si.Verb = "runas";
                Process.Start(si);
                Close();
            }
            catch (Win32Exception) { }      // said no at the uac prompt, carry on as we were
            catch (Exception ex) { Say("Could not restart elevated.\r\n\r\n" + ex.Message); }
        }

        static string Quote(string s)
        {
            return "\"" + (s.EndsWith("\\", StringComparison.Ordinal) ? s + "\\" : s) + "\"";
        }

        void ShowUnreadable()
        {
            if (scan == null) return;
            string[] list = unread.ToArray();
            using (Form f = new Form())
            {
                f.Text = "Folders that could not be read";
                f.StartPosition = FormStartPosition.CenterParent;
                f.ClientSize = new Size(760, 420);
                f.Font = Ui.Base;
                f.MinimizeBox = false;
                f.ShowInTaskbar = false;
                TextBox t = new TextBox();
                t.Multiline = true;
                t.ReadOnly = true;
                t.WordWrap = false;
                t.ScrollBars = ScrollBars.Both;
                t.Dock = DockStyle.Fill;
                t.BackColor = Ui.White;
                string note = notAdmin ? "Not running as administrator, so some of these may just need an admin window.\r\n\r\n" : "";
                t.Text = note + string.Join("\r\n", list)
                    + (unreadCount > list.Length ? "\r\n... and " + Fmt.Count(unreadCount - list.Length) + " more" : "");
                t.Select(0, 0);
                f.Controls.Add(t);
                f.ShowDialog(this);
            }
        }

        void Say(string message)
        {
            MessageBox.Show(this, message, "OOutSize", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    public static class Program
    {
        // the script reads this. a powershell window that already compiled an older copy keeps that copy,
        // and this is how the script can tell and say so. goes up every time the c# changes - the number in
        // the script's ;ast few lines has to match
        public const int Build = 5;

        // called from powershell. the window gets its own sta thread so it does not care how the host was started
        public static void Run(string path, string script, bool backupPrivilege, int maxFilesPerFolder, int threads,
            bool classicScan, bool readOnly)
        {
            Options o = new Options();
            o.Path = path ?? "";
            o.Script = script ?? "";
            o.Backup = backupPrivilege;
            o.PerFolder = maxFilesPerFolder;
            o.Threads = threads;
            o.Classic = classicScan;
            o.ReadOnly = readOnly;
            Thread t = new Thread(delegate() { Show(o); });
            t.SetApartmentState(ApartmentState.STA);
            t.Name = "ooutsize ui";
            t.Start();
            t.Join();
        }

        static void Show(Options o)
        {
            ThreadExceptionEventHandler oops = delegate(object s, ThreadExceptionEventArgs e)
            {
                MessageBox.Show(e.Exception.ToString(), "OOutSize hit a problem", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };
            try
            {
                Application.EnableVisualStyles();
                Application.ThreadException += oops;
                using (MainForm f = new MainForm(o)) Application.Run(f);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "OOutSize hit a problem", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { Application.ThreadException -= oops; }
        }
    }
}
'@

# compiled once per powershell session, and a session cannot swap a compiled copy for a newer one.
# so when an older build is already loaded in this window, say so instead of running the wrong code
$loaded = 'OOutSize.Program' -as [type]
if ($loaded) {
    $build = $loaded.GetField('Build')
    if (-not $build -or $build.GetValue($null) -ne 5) {
        throw 'A different version of OOutSize is already loaded in this PowerShell window. Open a new PowerShell window and run it again.'
    }
}
else {
    Write-Host 'Building OOutSize, this takes a second or two...'
    Add-Type -TypeDefinition $source -ReferencedAssemblies 'System.Windows.Forms', 'System.Drawing' -IgnoreWarnings
}

[OOutSize.Program]::Run($Path, [string]$PSCommandPath, [bool]$BackupPrivilege, $MaxFilesPerFolder, $Threads, [bool]$ClassicScan, [bool]$ReadOnly)
