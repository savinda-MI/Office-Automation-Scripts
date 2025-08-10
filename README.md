Office Automation Scripts

## Scripts in this Repository

- **DeleteOldFiles.vbs**  
  A VBScript that deletes files older than a specified number of days from a single folder and all its subfolders. Suitable for network shares or folders where more control and recursion are needed.

- **CleanupBackups.bat**  
  A batch script that deletes old files from multiple backup folders using the `forfiles` command. It quickly cleans specific folders and is ideal for straightforward folder cleanup on local or mapped drives.

- **DailyUpdateCopy.bat**  
  A batch script using `robocopy` to copy files modified within the last day from multiple source folders to corresponding destination folders. Ideal for automating daily file backups or synchronization tasks.

## Usage Notes

### DeleteOldFiles.vbs

Edit these lines before running the script:

    NumberOfDaysOld = 5    ' <-- Change this to the number of days to keep files
    strPath = "\\172.31.37.10\crm\CRM"   ' <-- Change this to your target folder path

### CleanupBackups.bat

Update the folder paths in each `forfiles` command to match your backup folders, for example:

    forfiles -p "D:\bck\S2" -s -m *.* /D -1 /C "cmd /c del @path"    REM <-- Change "D:\bck\S2" and other folder paths as needed

### DailyUpdateCopy.bat

Modify the source and destination folder paths for each `robocopy` line, for example:

    robocopy "Z:\Olax\Olax Automated Backup\Ambatenna" "Z:\Olax\New folder\Ambatenna" /s /maxage:1

- `/s` copies subdirectories excluding empty ones.  
- `/maxage:1` copies only files modified within the last 1 day.

You can add or remove lines as needed for different folders.

## Difference Between the Scripts

- `DeleteOldFiles.vbs`  
  Offers more control and recursive deletion through subfolders, ideal for network drives or complex folder structures.

- `CleanupBackups.bat`  
  Uses the Windows `forfiles` command to quickly delete files in specified folders, suitable for simple and fast cleanup on local or mapped drives.

- `DailyUpdateCopy.bat`  
  Copies recently modified files daily from source to destination folders, useful for backups and sync.

---

## Author

Savinda Missaka  
[LinkedIn](https://www.linkedin.com/in/savinda-missaka-52b49425a)  
savinda.missaka@gmail.com
