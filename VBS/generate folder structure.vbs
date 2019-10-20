Dim fso1, fso2 , flds, subfld, fld , file , foldername, outputfilename
'use your first folder name here:
foldername="C:\Documents and Settings\dan_petran\Desktop\CRM OD file attachments"
outputfilename="C:\Documents and Settings\dan_petran\Desktop\CRM OD file attachments\dirlist.xls"
'create file system object for folder names
Set fso1 = CreateObject("Scripting.FileSystemObject")
Set flds = fso1.Getfolder(foldername)
Set subfld = flds.SubFolders
'create file system object for text file
Set fso2 = CreateObject("Scripting.FileSystemObject")
Set file = fso2.OpenTextFile(outputfilename, 2, True)

'get each folder name and save in text file comma delimited
For Each fld in subfld
   file.Write fld.name & chr(10)
Next
file.WriteLine ""
file.Close
set fso1=Nothing
set fso2=Nothing
