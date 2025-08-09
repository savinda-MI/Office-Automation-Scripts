NumberOfDaysOld = 5
strPath = "\\172.31.37.10\crm\CRM"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strPath)
Set colSubfolders = objFolder.Subfolders
Set colFiles = objFolder.Files

For Each objFile in colFiles
If objFile.DateLastModified < (Date() - NumberOfDaysOld) Then
objFile.Delete
End If
Next

For Each objSubfolder in colSubfolders
Set colFiles = objSubfolder.Files
For Each objFile in colFiles
If objFile.DateLastModified < (Date() - NumberOfDaysOld) Then
objFile.Delete
End If
Next

Next