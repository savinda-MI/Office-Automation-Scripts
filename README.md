Office Automation Scripts

## Scripts in this Repository

- **DeleteOldFiles.vbs**  
  A VBScript that deletes files older than a specified number of days from a single folder and all its subfolders. Suitable for network shares or folders where more control and recursion are needed.

- **CleanupBackups.bat**  
  A batch script that deletes old files from multiple backup folders using the `forfiles` command. It quickly cleans specific folders and is ideal for straightforward folder cleanup on local or mapped drives.

- **CleanupBackups_New.bat**  
  An improved batch script that can delete old files only if the total file count in a folder exceeds a set threshold. This prevents accidental deletion in cases where a folder already has few files. It’s useful for log or backup directories where you want to maintain a minimum number of files.

## Usage Notes

### DeleteOldFiles.vbs

Edit these lines before running the script:

    NumberOfDaysOld = 5    ' <-- Change this to the number of days to keep files
    strPath = "\\172.31.37.10\crm\CRM"   ' <-- Change this to your target folder path

### CleanupBackups.bat

Update the folder paths in each `forfiles` command to match your backup folders, for example:

    forfiles -p "D:\bck\S2" -s -m *.* /D -1 /C "cmd /c del @path"    REM <-- Change "D:\bck\S2" and other folder paths as needed

### CleanupBackups_New.bat

Edit these variables before running:

    set "folder=D:\bck\S2"         REM <-- Target folder path
    set "maxCount=50"              REM <-- Minimum number of files to keep before deletion starts
    set "daysOld=1"                REM <-- Files older than this number of days will be deleted

This script will first count the number of files in the folder. If the count is greater than `maxCount`, it will delete files older than `daysOld` days. Otherwise, no deletion occurs.

## Difference Between the Scripts

- `DeleteOldFiles.vbs`  
  Offers more control and recursive deletion through subfolders, making it ideal for network drives or complex folder structures.

- `CleanupBackups.bat`  
  Uses the Windows `forfiles` command to quickly delete files in specified folders, suitable for simple and fast cleanup on local or mapped drives.

- `CleanupBackups_New.bat`  
  Adds a safety mechanism by checking file count before deletion. Best for scenarios where you want to keep a minimum number of files in a folder while still removing older ones.

---

## Author

Savinda Missaka  
[LinkedIn](https://www.linkedin.com/in/savinda-missaka-52b49425a)  
savinda.missaka@gmail.com
