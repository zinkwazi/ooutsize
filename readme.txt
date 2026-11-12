OOutSize
========

OOutSize shows what is using the space on a drive or folder. A folder tree
with size bars sits next to a treemap you can drill into.

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

When that is the case a yellow strip across the top of the window says so,
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
- free: free space on the volume. The tile is green, yellow under 20 percent
  or red under 10 percent

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
- Right click for Open in Explorer, Copy path and Copy listing
- Click a file type under the map to keep its colour and fade out the rest.
  Click it again to switch back
- F5 scans again
- Esc stops a scan and shows what it has so far


System files
------------

Some of the biggest things on a Windows drive must not be deleted by hand.
Hover over one and OOutSize says why and what to do instead.

Striped, because they belong to Windows

- Windows, including WinSxS, Installer and SoftwareDistribution
- pagefile.sys, swapfile.sys and hiberfil.sys
- System Volume Information, Recovery, Boot and EFI

Not striped, but the same warning shows when you hover

- Program Files and Program Files (x86)
- ProgramData\Microsoft and ProgramData\Package Cache
- $Recycle.Bin

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


Is it safe to run
-----------------

- It only lists folders: names, sizes and dates
- It never opens, changes or deletes a file. There is no delete button
- It writes nothing to the drive it scans. The only thing written is a
  small temporary file in your TEMP folder while Windows compiles the tool
- Security software may log PowerShell starting csc.exe at launch. That is
  the compile step and it is expected


Options
-------

-Path <drive or folder>
    What to scan, for example C:\ or D:\Shares or \\server\share

-BackupPrivilege
    From an admin window, lets the scan list folders that even
    administrators are normally locked out of, such as System Volume
    Information and other people's profiles. Off unless you ask for it

-MaxFilesPerFolder <number>
    How many files in one folder get a row of their own. Default 5000.
    0 means no limit

Example

    powershell -ExecutionPolicy Bypass -File C:\Temp\OOutSize.ps1 -Path D:\Shares -BackupPrivilege

Full help

    Get-Help C:\Temp\OOutSize.ps1 -Full


If something goes wrong
-----------------------

"running scripts is disabled on this system"
    Start it with the powershell -ExecutionPolicy Bypass -File line above

"PowerShell is running in ConstrainedLanguage mode"
    AppLocker or WDAC is restricting PowerShell on this computer and the
    tool cannot start there

Red compiler errors when it starts
    Copy the text of the first error and send it to whoever gave you
    this file

The scan is slow
    A large drive with millions of files can take several minutes. The
    progress screen shows what it has found so far. Esc stops it early

The numbers look too small
    Check for the yellow strip. Without admin rights, folders the account
    cannot read are missing from the totals
