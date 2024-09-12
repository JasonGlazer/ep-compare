#tag Class
Protected Class App
Inherits Application
	#tag Event
		Sub Activate()
		  'check if any of the files changed
		  'dim changedFiles(-1) as String
		  'dim curFile as FolderItem
		  'dim i as integer
		  'for i = 0 to ActiveFileList.Ubound
		  'curFile = new FolderItem(ActiveFileList(i))
		  'if ActiveFileTimeStamp(i)< curFile.ModificationDate then
		  'changedFiles.Append ActiveFileList(i)
		  'end if
		  'next i
		  'if changedFiles.Ubound>=0 then
		  'MsgBox "Files changed: " + Join(changedFiles,", ") + " Files will be reloaded"
		  'Call ClearMainArrays
		  'Call LoadAllFiles
		  'Call MainWindow.PopulateList
		  'end if
		End Sub
	#tag EndEvent

	#tag Event
		Sub Close()
		  call WriteOldFileList
		  
		End Sub
	#tag EndEvent

	#tag Event
		Sub Open()
		  dim x as integer=1000
		  CDBaseChartMBS.setLicenseCode app.MyName,201006, 843479*x+311,app.MBSserial
		  CALL ReadGraphHints
		  CALL ReadOldFileList
		  
		  
		  'Dim currentFolder as FolderItem
		  'currentFolder=GetFolderItem("")
		  'MsgBox currentFolder.AbsolutePath
		  
		  
		  
		  
		  
		End Sub
	#tag EndEvent


	#tag Constant, Name = kEditClear, Type = String, Dynamic = False, Default = \"&Delete", Scope = Public
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"&Delete"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"&Delete"
	#tag EndConstant

	#tag Constant, Name = kFileQuit, Type = String, Dynamic = False, Default = \"&Quit", Scope = Public
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"E&xit"
	#tag EndConstant

	#tag Constant, Name = kFileQuitShortcut, Type = String, Dynamic = False, Default = \"", Scope = Public
		#Tag Instance, Platform = Mac OS, Language = Default, Definition  = \"Cmd+Q"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"Ctrl+Q"
	#tag EndConstant

	#tag Constant, Name = MBSserial, Type = Double, Dynamic = False, Default = \"1203213811 ", Scope = Public
	#tag EndConstant

	#tag Constant, Name = MyName, Type = String, Dynamic = False, Default = \"Jason Glazer", Scope = Public
	#tag EndConstant


	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
