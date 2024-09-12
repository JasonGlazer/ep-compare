#tag Window
Begin Window AboutWindow
   BackColor       =   &cFFFFFF00
   Backdrop        =   0
   CloseButton     =   False
   Compatibility   =   ""
   Composite       =   False
   Frame           =   0
   FullScreen      =   False
   FullScreenButton=   False
   HasBackColor    =   False
   Height          =   411
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
   Placement       =   0
   Resizeable      =   False
   Title           =   "About EP-Compare"
   Visible         =   True
   Width           =   553
   Begin Label StaticText1
      AutoDeactivate  =   True
      Bold            =   True
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   46
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   20
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Multiline       =   False
      Scope           =   0
      Selectable      =   False
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "EP-Compare"
      TextAlign       =   0
      TextColor       =   &c000000
      TextFont        =   "System"
      TextSize        =   32.0
      TextUnit        =   0
      Top             =   14
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   224
   End
   Begin Label StaticText2
      AutoDeactivate  =   True
      Bold            =   False
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   29
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   20
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Multiline       =   False
      Scope           =   0
      Selectable      =   False
      TabIndex        =   2
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "Version 0.4"
      TextAlign       =   0
      TextColor       =   &c000000
      TextFont        =   "System"
      TextSize        =   18.0
      TextUnit        =   0
      Top             =   62
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   235
   End
   Begin Listbox AboutList
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
      EnableDragReorder=   False
      GridLinesHorizontal=   0
      GridLinesVertical=   0
      HasHeading      =   False
      HeadingIndex    =   -1
      Height          =   262
      HelpTag         =   ""
      Hierarchical    =   False
      Index           =   -2147483648
      InitialParent   =   ""
      InitialValue    =   ""
      Italic          =   False
      Left            =   22
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      RequiresSelection=   False
      Scope           =   0
      ScrollbarHorizontal=   False
      ScrollBarVertical=   True
      SelectionType   =   0
      ShowDropIndicator=   False
      TabIndex        =   3
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   95
      Transparent     =   True
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   511
      _ScrollOffset   =   0
      _ScrollWidth    =   -1
   End
   Begin BevelButton CloseAboutButton
      AcceptFocus     =   False
      AutoDeactivate  =   True
      BackColor       =   &c000000
      Bevel           =   0
      Bold            =   False
      ButtonType      =   0
      Caption         =   "Close"
      CaptionAlign    =   3
      CaptionDelta    =   0
      CaptionPlacement=   1
      Enabled         =   True
      HasBackColor    =   False
      HasMenu         =   0
      Height          =   22
      HelpTag         =   ""
      Icon            =   0
      IconAlign       =   0
      IconDX          =   0
      IconDY          =   0
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   398
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MenuValue       =   0
      Scope           =   0
      TabIndex        =   4
      TabPanelIndex   =   0
      TabStop         =   True
      TextColor       =   &c000000
      TextFont        =   "System"
      TextSize        =   0.0
      TextUnit        =   0
      Top             =   369
      Transparent     =   True
      Underline       =   False
      Value           =   False
      Visible         =   True
      Width           =   135
   End
End
#tag EndWindow

#tag WindowCode
	#tag Event
		Sub Activate()
		  AboutList.AddRow "Copyright (c) 2009-2010 GARD Analytics, Inc.  All rights reserved."
		  AboutList.AddRow "                                        "
		  AboutList.AddRow "NOTICE: The U.S. Government is granted for itself and others acting on its behalf a paid-up,"
		  AboutList.AddRow "nonexclusive, irrevocable, worldwide license in this  data to reproduce, prepare derivative "
		  AboutList.AddRow "works, and perform publicly and display publicly. Beginning five (5) years after permission"
		  AboutList.AddRow "to assert copyright is granted, subject to two possible five year renewals, the U.S. Government"
		  AboutList.AddRow "is granted for itself and others acting on its behalf a paid-up, non-exclusive ,irrevocable"
		  AboutList.AddRow "worldwide license in this data to reproduce, prepare derivative works, distribute copies to"
		  AboutList.AddRow "the public, perform publicly and display publicly,and to permit others to do so"
		  AboutList.AddRow "                                        "
		  AboutList.AddRow "TRADEMARKS: EnergyPlus, DOE-2.1E, DOE-2, and DOE are trademarks of the US "
		  AboutList.AddRow "Department of Energy.                   "
		  AboutList.AddRow "                                        "
		  AboutList.AddRow "DISCLAIMER OF WARRANTY AND LIMITATION OF LIABILITY: THIS SOFTWARE IS PROVIDED 'AS"
		  AboutList.AddRow "IS' WITHOUT WARRANTY OF ANY KIND. NEITHER GARD ANALYTICS, THE DEPARTMENT  OF"
		  AboutList.AddRow "ENERGY, THE US GOVERNMENT, THEIR LICENSORS, OR ANY PERSON OR ORGANIZATION"
		  AboutList.AddRow "ACTING ON BEHALF OF ANY OF THEM:"
		  AboutList.AddRow "                                        "
		  AboutList.AddRow "A.  MAKE ANY WARRANTY OR REPRESENTATION WHATSOEVER, EXPRESS OR IMPLIED, WITH"
		  AboutList.AddRow "RESPECT TO ENERGYPLUS OR ANY DERIVATIVE WORKS THEREOF, INCLUDING WITHOUT"
		  AboutList.AddRow "LIMITATION WARRANTIES OF MERCHANTABILITY, WARRANTIES OF FITNESS  FOR A"
		  AboutList.AddRow "PARTICULAR PURPOSE, OR WARRANTIES OR REPRESENTATIONS REGARDING THE USE, OR"
		  AboutList.AddRow "THE RESULTS OF THE USE OF ENERGYPLUS OR DERIVATIVE WORKS THEREOF IN TERMS OF"
		  AboutList.AddRow "CORRECTNESS, ACCURACY, RELIABILITY, CURRENTNESS, OR OTHERWISE. THE ENTIRE"
		  AboutList.AddRow "RISK AS TO THE RESULTS AND PERFORMANCE OF THE LICENSED SOFTWARE IS ASSUMED BY"
		  AboutList.AddRow "THE LICENSEE.                           "
		  AboutList.AddRow "                                        "
		  AboutList.AddRow "B.  MAKE ANY REPRESENTATION OR WARRANTY THAT ENERGYPLUS OR DERIVATIVE WORKS"
		  AboutList.AddRow "THEREOF WILL NOT INFRINGE ANY COPYRIGHT OR OTHER PROPRIETARY RIGHT."
		  AboutList.AddRow "                                        "
		  AboutList.AddRow "C.  ASSUME ANY LIABILITY WHATSOEVER WITH RESPECT TO ANY USE OF ENERGYPLUS,"
		  AboutList.AddRow "DERIVATIVE WORKS THEREOF, OR ANY PORTION THEREOF OR WITH RESPECT TO ANY DAMAGES"
		  AboutList.AddRow "WHICH MAY RESULT FROM SUCH USE.         "
		  AboutList.AddRow "                                        "
		  AboutList.AddRow "DISCLAIMER OF ENDORSEMENT: Reference herein to any specific commercial products,"
		  AboutList.AddRow "process, or service by trade name, trademark, manufacturer, or otherwise, does not"
		  AboutList.AddRow "necessarily constitute or imply its endorsement, recommendation, or favoring by the"
		  AboutList.AddRow "United States Government or GARD Analytics, Inc."
		  
		End Sub
	#tag EndEvent


#tag EndWindowCode

#tag Events CloseAboutButton
	#tag Event
		Sub Action()
		  AboutWindow.Close
		End Sub
	#tag EndEvent
#tag EndEvents
