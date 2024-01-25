Dim SysFile, SysFolder, FilesInFolder
'Give you folder path here
'NOTE: path must end with \
CurrentPath="D:\Download\"

Set SysFile= CreateObject("Scripting.FileSystemObject")
Set SysFolder = SysFile.getfolder(CurrentPath)

'loop through each files in folder
For Each FilesInFolder In SysFolder.Files
  'get file name
	FileName = FilesInFolder.Name
  'get file extension of file
	FileExtension = SysFile.GetExtensionName(CurrentPath&FileName)
  'trim file extension
	FileExtentionName = Left(FileName, Len(FileName) - Len(FileExtension) -1)
  'msgbox FileExtentionName 

  'create folder for extension
	extensionfolder = CurrentPath & FileExtension
  'msgbox extensionfolder
	If Not SysFile.FolderExists(extensionfolder) Then
		Set extensionfolder = SysFile.CreateFolder(extensionfolder)
	End If
  'move file 
  source=CurrentPath & FileName
	ModifiedDate = FilesInFolder.DateLastModified
	ModifiedDate1 = FormatDateTime(ModifiedDate,2)
  'msgbox ModifiedDate1
	destination =CurrentPath & FileExtension & "\" & FileExtentionName & " " & ModifiedDate1 & "." & FileExtension
  'msgbox source
  'msgbox destination
	SysFile.MoveFile source, destination
Next
