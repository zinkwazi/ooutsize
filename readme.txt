OOutSize
========

OOutSize shows what is using the space on a drive or folder. A folder tree
with size bars sits next to a treemap you can drill into. You can delete
straight from it, with a confirm box every time.

OOutSize is open source so you can audit the safety before use.
It is one PowerShell file. There is nothing to install.

https://github.com/zinkwazi/ooutsize

How to run it
-------------

- Copy OOutSize.ps1 to the computer you want to look at, for example C:\Temp
- Open PowerShell as administrator
- Run this, with the drive or folder you want at the end

    powershell -ExecutionPolicy Bypass -File C:\Temp\OOutSize.ps1 C:\

Leave the path off to pick from a list of drives.

You can also right click the file and choose Run with PowerShell. That
starts it without admin rights, see the next section.

The first few seconds are Windows compiling the tool. Then the window opens.


What you need
-------------

- Windows 10, Windows 11 or Windows Server 2016 and later
- Windows PowerShell and the .NET Framework, which are both built in

If you start it from PowerShell 7 it hands itself over to Windows PowerShell.


Admin rights
------------

OOutSize runs without admin rights, but then it can only see what your
account can read. Folders it cannot get into are left out, so the totals
come up short.

When that is the case an orange strip across the top of the window says so,
with a Restart as admin button. The PowerShell window prints the same warning.

Being signed in as an administrator is not enough on its own. The PowerShell
window itself has to be opened with Run as administrator.


Reading the window
------------------

Top row of tiles

- scanned: total size of everything found
- files and folders: how many
- in use but not visible: on a whole drive scan, space the drive says is
  used that the scan could not see. This is usually shadow copies, folders
  it was not allowed into or file system overhead
- could not be read: folders that were skipped. Click the tile for the list
- free: free space on the volume. Under 20 percent the label turns red with
  a down arrow and the word low. Under 10 percent it is bold as well

Under the path it says how long the scan took and on how many threads.

Left side, the tree

- Folders and files, biggest first
- Share of parent shows how much of the folder above each row takes up
- Newest is the date of the most recent file inside. An orange tag means
  nothing in there has changed for over a year

Right side, the map

- Every block is a file or folder and its area is its size
- Each top level folder gets its own soft colour, with its contents laid
  out inside
- The seven file types using the most space each get a colour. The key is
  under the map. Everything else is pale orange
- Striped blocks are Windows system files, see below


Mouse and keys
--------------

- Click a folder in the tree to show it in the map
- Arrow keys move around the tree
- Double click a block in the map to zoom in
- Backspace or the Up button goes back out
- Hover over a block for its name, size and date
- Right click for Open in Explorer, Copy path, Copy listing, Scan again,
  Move to Recycle Bin and Delete permanently
- Del moves the selected row to the Recycle Bin. Shift+Del deletes it
  permanently. Both ask first
- Point at a file type under the map to fade out the rest. Click it to keep
  it that way and click again to switch back
- F5 scans again
- Esc stops a scan and shows what it has so far


Deleting
--------

Right click a file or folder and pick Move to Recycle Bin or Delete
permanently. Nothing happens until you answer Yes in the box that follows.
The box shows the full path and the size from the scan. No is the default
button.

- The delete is handed to Windows itself, the same code Explorer uses. Its
  usual windows may appear: progress, a file in use, a request for admin
  rights
- Afterwards the tree, the map and the tiles update. No full rescan needed
- Recycle Bin: the space is not freed until the bin is emptied. On a whole
  drive scan you will see the item turn up under $Recycle.Bin
- Network drives have no Recycle Bin. Windows asks before it deletes for
  good instead
- If only part of a folder could be deleted, OOutSize scans that folder
  again and shows what is left

What OOutSize will not delete

- Anything striped, see System files below
- These whole folders on the Windows drive: Users, ProgramData,
  ProgramData\Microsoft, Program Files and Program Files (x86). Things
  inside them can be deleted
- The folder the scan started in
- A "smaller files" row, which is many files rolled into one
- Anything with a path over 259 characters or with something that deep
  inside it
- Anything with a name ending in a dot or a space

The Windows delete function cannot be trusted with those last two.

What gets a sterner question first

- Anything inside Program Files, ProgramData\Microsoft,
  ProgramData\Package Cache or $Recycle.Bin
- Windows upgrade leftovers such as Windows.old
- A user profile folder

Scan again on the right click menu rescans one folder and leaves the rest
of the tree alone. Use it after cleaning something up outside OOutSize.

Start OOutSize with -ReadOnly to take the delete options away.


System files
------------

Some of the biggest things on a Windows drive must not be deleted by hand.
Hover over one and OOutSize says why and what to do instead.

Striped, because they belong to Windows. Delete is switched off for these

- Windows, including WinSxS, Installer and SoftwareDistribution
- pagefile.sys, swapfile.sys and hiberfil.sys
- System Volume Information, Recovery, Boot and EFI

Inside Windows these are not striped, because they only hold temporary
files, logs and crash dumps: Temp, Logs, Minidump, LiveKernelReports and
MEMORY.DMP.

Not striped, but a warning shows when you hover or delete

- Program Files and Program Files (x86)
- ProgramData\Microsoft and ProgramData\Package Cache
- $Recycle.Bin
- Windows.old, $Windows.~BT and $Windows.~WS

This is a fixed list matched by path. It is a guide, not a guarantee.
Something with no warning is not automatically safe to delete.


What the sizes mean
-------------------

- Sizes match what Explorer calls Size
- Sparse, compressed and cloud placeholder files are the exception. They
  count as the space they really take on disk
- Junctions and symbolic links are not followed
- Hard linked files count once per name, the same as Explorer. This makes
  C:\Windows\WinSxS read higher than it really is
- A folder lists its 5000 biggest files one by one. Any more are rolled
  into a single "smaller files" row. Nothing is left out of the totals


Speed
-----

OOutSize walks the folders on several threads at once. The first scan
after a restart is the slowest one, because Windows has to read the folder
information from the disk. A second scan of the same drive is much quicker.

Tools such as TreeSize and WizTree can be faster on a local NTFS drive
because they read the Master File Table straight off the disk instead of
walking the folders. OOutSize does not do that. It asks Windows for folder
listings the normal way, which also works on network shares and other
file systems.


Is it safe to run
-----------------

- Scanning only lists folders: names, sizes and dates. It never opens or
  changes a file
- Nothing is deleted unless you choose one of the two delete options and
  then answer Yes. -ReadOnly removes those options
- Apart from a delete you ask for, it writes nothing to the drive it scans.
  The only other thing written is a small temporary file in your TEMP
  folder while Windows compiles the tool
- Security software may log PowerShell starting csc.exe at launch. That is
  the compile step and it is expected


Options
-------

-Path <drive or folder>
    What to scan, for example C:\ or D:\Shares or \\server\share

-ReadOnly
    Leave out the delete options, for a window that can only look

-BackupPrivilege
    From an admin window, lets the scan list folders that even
    administrators are normally locked out of, such as System Volume
    Information and other people's profiles. Off unless you ask for it

-Threads <number>
    How many threads list folders at once. The default is twice the number
    of processors, at least 4 and at most 32. Use a low number such as 2
    to go easy on a busy file server

-MaxFilesPerFolder <number>
    How many files in one folder get a row of their own. Default 5000.
    0 means no limit

-ClassicScan
    List folders the way the first version did. It is slower. Try it if
    the totals look wrong on an unusual file system or network share

Examples

    powershell -ExecutionPolicy Bypass -File C:\Temp\OOutSize.ps1 -Path D:\Shares -BackupPrivilege

    powershell -ExecutionPolicy Bypass -File C:\Temp\OOutSize.ps1 -Path \\server\share -ReadOnly

Full help

    Get-Help C:\Temp\OOutSize.ps1 -Full


If something goes wrong
-----------------------

"running scripts is disabled on this system"
    Start it with the powershell -ExecutionPolicy Bypass -File line above

"PowerShell is running in ConstrainedLanguage mode"
    AppLocker or WDAC is restricting PowerShell on this computer and the
    tool cannot start there

"A different version of OOutSize is already loaded"
    That PowerShell window ran an older copy earlier. Open a new
    PowerShell window and run it again

Red compiler errors when it starts
    Copy the text of the first error and send it to whoever gave you
    this file

The scan is slow
    A large drive with millions of files can take several minutes. The
    first scan after a restart is the slowest. The progress screen shows
    what it has found so far. Esc stops it early

The numbers look too small
    Check for the orange strip. Without admin rights, folders the account
    cannot read are missing from the totals

There is no delete on the right click menu
    Either it was started with -ReadOnly or it is an older copy of the
    file. The current one shows the thread count next to the scan time

A delete did nothing
    Windows refused it, usually because the file is in use or the account
    is not allowed. The status line at the bottom says what happened
