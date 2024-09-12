#tag Window
Begin Window ManageFilesWindow
   BackColor       =   &cFFFFFF00
   Backdrop        =   0
   CloseButton     =   True
   Compatibility   =   ""
   Composite       =   False
   Frame           =   1
   FullScreen      =   False
   FullScreenButton=   False
   HasBackColor    =   False
   Height          =   227
   ImplicitInstance=   True
   LiveResize      =   False
   MacProcID       =   0
   MaxHeight       =   32000
   MaximizeButton  =   False
   MaxWidth        =   32000
   MenuBar         =   0
   MenuBarVisible  =   True
   MinHeight       =   64
   MinimizeButton  =   False
   MinWidth        =   64
   Placement       =   1
   Resizeable      =   True
   Title           =   "Manage Files"
   Visible         =   True
   Width           =   1051
   Begin Listbox ListOfFiles
      AutoDeactivate  =   True
      AutoHideScrollbars=   True
      Bold            =   False
      Border          =   True
      ColumnCount     =   1
      ColumnsResizable=   False
      ColumnWidths    =   ""
      DataField       =   ""
      DataSource      =   ""
      DefaultRowHeight=   -1
      Enabled         =   True
      EnableDrag      =   False
      EnableDragReorder=   True
      GridLinesHorizontal=   0
      GridLinesVertical=   0
      HasHeading      =   False
      HeadingIndex    =   -1
      Height          =   162
      HelpTag         =   ""
      Hierarchical    =   False
      Index           =   -2147483648
      InitialParent   =   ""
      InitialValue    =   ""
      Italic          =   False
      Left            =   20
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      RequiresSelection=   False
      Scope           =   0
      ScrollbarHorizontal=   False
      ScrollBarVertical=   True
      SelectionType   =   0
      ShowDropIndicator=   False
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   14
      Transparent     =   True
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   878
      _ScrollOffset   =   0
      _ScrollWidth    =   -1
   End
   Begin PushButton btnAddFile
      AutoDeactivate  =   True
      Bold            =   False
      ButtonStyle     =   "0"
      Cancel          =   False
      Caption         =   "Add File"
      Default         =   False
      Enabled         =   True
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   904
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      Scope           =   0
      TabIndex        =   4
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   14
      Transparent     =   True
      Underline       =   False
      Visible         =   True
      Width           =   137
   End
   Begin PushButton btnClose
      AutoDeactivate  =   True
      Bold            =   False
      ButtonStyle     =   "0"
      Cancel          =   False
      Caption         =   "Close"
      Default         =   True
      Enabled         =   True
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   904
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   False
      Scope           =   0
      TabIndex        =   5
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   187
      Transparent     =   True
      Underline       =   False
      Visible         =   True
      Width           =   137
   End
   Begin PushButton btnAddDirectory
      AutoDeactivate  =   True
      Bold            =   False
      ButtonStyle     =   "0"
      Cancel          =   False
      Caption         =   "Add Directory"
      Default         =   True
      Enabled         =   True
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   904
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      Scope           =   0
      TabIndex        =   6
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   48
      Transparent     =   True
      Underline       =   False
      Visible         =   True
      Width           =   137
   End
   Begin PushButton btnRemoveAll
      AutoDeactivate  =   True
      Bold            =   False
      ButtonStyle     =   "0"
      Cancel          =   False
      Caption         =   "Remove All"
      Default         =   False
      Enabled         =   True
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   904
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      Scope           =   0
      TabIndex        =   7
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   134
      Transparent     =   True
      Underline       =   False
      Visible         =   True
      Width           =   137
   End
   Begin PushButton btnRemoveOne
      AutoDeactivate  =   True
      Bold            =   False
      ButtonStyle     =   "0"
      Cancel          =   False
      Caption         =   "Remove"
      Default         =   False
      Enabled         =   True
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   904
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      Scope           =   0
      TabIndex        =   8
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   100
      Transparent     =   True
      Underline       =   False
      Visible         =   True
      Width           =   137
   End
   Begin Label StaticText1
      AutoDeactivate  =   True
      Bold            =   False
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   19
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   True
      Left            =   20
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   False
      Multiline       =   False
      Scope           =   0
      Selectable      =   False
      TabIndex        =   9
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "The top file in the list is used as the baseline file. Drag files to change order.\r\n"
      TextAlign       =   0
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   188
      Transparent     =   True
      Underline       =   False
      Visible         =   True
      Width           =   497
   End
End
#tag EndWindow

#tag WindowCode
	#tag Event
		Sub Close()
		  dim i as Integer
		  
		  redim ActiveFileList(-1)
		  for i = 0 to ListOfFiles.ListCount - 1
		    ActiveFileList.Append ListOfFiles.List(i)
		  next i
		  Call UpdateActiveFileTimeStamps
		End Sub
	#tag EndEvent

	#tag Event
		Sub Open()
		  dim i as integer
		  
		  ListOfFiles.DeleteAllRows
		  for i = 0 to ubound(ActiveFileList)
		    ListOfFiles.AddRow  ActiveFileList(i)
		  next i
		End Sub
	#tag EndEvent


#tag EndWindowCode

#tag Events btnAddFile
	#tag Event
		Sub Action()
		  Dim FileToAdd as FolderItem
		  dim found as Boolean
		  dim i as integer
		  
		  found = false
		  FileToAdd = GetOpenFolderItem(FileTypeHTML.All)
		  If FileToAdd <> Nil then
		    for i = 1 to ListOfFiles.ListCount
		      if FileToAdd.AbsolutePath = ListofFiles.List(i - 1) then
		        found = true
		      end if
		    next i
		    if  found then
		      MsgBox "File already on list: " + FileToAdd.AbsolutePath
		    else
		      ListOfFiles.AddRow FileToAdd.AbsolutePath
		    end if
		  End if
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnClose
	#tag Event
		Sub Action()
		  ManageFilesWindow.Close
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnAddDirectory
	#tag Event
		Sub Action()
		  Dim FolderSelected as FolderItem
		  dim FileToAdd as FolderItem
		  dim found as Boolean
		  dim i as integer
		  dim j as integer
		  
		  FolderSelected = SelectFolder
		  If FolderSelected <> Nil then
		    if FolderSelected.Directory then
		      for i = 1 to FolderSelected.Count
		        found = false
		        FileToAdd = FolderSelected.Item(i)
		        if FileToAdd.Type = FileTypeHTML.TextHtml.Name then
		          for j = 1  to ListOfFiles.ListCount
		            if FileToAdd.AbsolutePath = ListofFiles.List(j - 1) then
		              found = true
		            end if
		          next j
		          if  not found then
		            ListOfFiles.AddRow FileToAdd.AbsolutePath
		          end if
		        end if
		      next i
		    End if
		  end if
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnRemoveAll
	#tag Event
		Sub Action()
		  ListOfFiles.DeleteAllRows
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnRemoveOne
	#tag Event
		Sub Action()
		  If listOfFiles.ListIndex >=0 and ListofFiles.ListIndex < ListOfFiles.ListCount then
		    ListOfFiles.RemoveRow ListOfFiles.ListIndex
		  End If
		End Sub
	#tag EndEvent
#tag EndEvents
