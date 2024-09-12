#tag Window
Begin Window MainWindow
   BackColor       =   &cFFFFFF00
   Backdrop        =   0
   CloseButton     =   True
   Compatibility   =   ""
   Composite       =   False
   Frame           =   0
   FullScreen      =   False
   FullScreenButton=   False
   HasBackColor    =   False
   Height          =   645
   ImplicitInstance=   True
   LiveResize      =   True
   MacProcID       =   0
   MaxHeight       =   32000
   MaximizeButton  =   True
   MaxWidth        =   32000
   MenuBar         =   924303359
   MenuBarVisible  =   True
   MinHeight       =   600
   MinimizeButton  =   True
   MinWidth        =   800
   Placement       =   2
   Resizeable      =   True
   Title           =   "EP-Compare"
   Visible         =   True
   Width           =   1094
   Begin Canvas PlotCanvas
      AcceptFocus     =   False
      AcceptTabs      =   False
      AutoDeactivate  =   True
      Backdrop        =   0
      DoubleBuffer    =   False
      Enabled         =   True
      EraseBackground =   True
      Height          =   420
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   7
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Scope           =   0
      TabIndex        =   11
      TabPanelIndex   =   0
      TabStop         =   True
      Top             =   214
      Transparent     =   True
      UseFocusRing    =   True
      Visible         =   True
      Width           =   1080
      Begin Listbox lstResults
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
         Height          =   413
         HelpTag         =   ""
         Hierarchical    =   False
         Index           =   -2147483648
         InitialParent   =   "PlotCanvas"
         InitialValue    =   "test value\r\n"
         Italic          =   False
         Left            =   13
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
         Top             =   221
         Transparent     =   False
         Underline       =   False
         UseFocusRing    =   True
         Visible         =   True
         Width           =   1073
         _ScrollOffset   =   0
         _ScrollWidth    =   -1
      End
   End
   Begin tbMain tbMain1
      Enabled         =   True
      Index           =   -2147483648
      InitialParent   =   ""
      LockedInPosition=   False
      Scope           =   0
      TabPanelIndex   =   0
      Visible         =   True
   End
   Begin TreeView TreeViewGraphs
      AutoDeactivate  =   True
      BackColor       =   &cFFFFFF00
      ColumnCount     =   1
      DragReceiveBehavior=   1
      Enabled         =   True
      HasBorder       =   False
      HasHeader       =   False
      HasInactiveSelectionColor=   False
      HasSelectionColor=   False
      Height          =   200
      HelpTag         =   ""
      InactiveSelectionColor=   &cEBEBEB00
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   7
      LinuxDrawTreeLines=   False
      LinuxExpanderStyle=   0
      LinuxHighlightFullRow=   True
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      MacDrawTreeLines=   False
      MacExpanderStyle=   0
      MacHighlightFullRow=   True
      MultiSelection  =   False
      NodeEvenColor   =   &cFFFFFF00
      NodeHeight      =   18
      NodeOddColor    =   &cFFFFFF00
      QuartzShading   =   False
      Scope           =   0
      SelectionColor  =   &c316AC500
      SelectionSeparator=   0
      TabIndex        =   13
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   7.0
      TextUnit        =   0
      Top             =   9
      UseFocusRing    =   True
      Visible         =   True
      Width           =   1080
      WinDrawTreeLines=   True
      WinHighlightFullRow=   True
   End
End
#tag EndWindow

#tag WindowCode
	#tag Event
		Sub Open()
		  lstResults.Visible=false
		  TreeViewGraphs.TextSize=12
		  'call drawChart
		  if ActiveFileList.Ubound >= 0 then
		    Call ClearMainArrays
		    Call LoadAllFiles
		    Call PopulateList
		  end if
		  'remove from view the unimplemented tool bar buttons
		  tbMain1.Remove 9
		  tbMain1.Remove 8
		  tbMain1.Remove 7
		  tbMain1.Remove 6
		  tbMain1.Remove 5
		  tbMain1.Remove 4
		  tbMain1.Remove 3
		  'tbMain1.Remove 2
		  tbMain1.Remove 1
		End Sub
	#tag EndEvent

	#tag Event
		Sub Resized()
		  call drawChart
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h0
		Sub DecodeGraphName(codeIn as string, ByRef hintOut as integer, ByRef tableNameOut as string, ByRef forOut as string, ByRef graphNameOut as string, ByRef headOneOut as string, ByRef headTwoOut as string, ByRef kindGraphOut as String)
		  'str(iHint) + " | " + ghCombinedName(iHint) + " | " + uniqueFors(nFor) + " | " + specificGraphNames(mSpGrph)
		  
		  dim phrase(-1) as string
		  dim i as integer
		  dim locFirstDashes as integer
		  dim locSecondDashes as integer
		  
		  
		  phrase = split(codeIn, "|")
		  for i  = 0 to phrase.Ubound
		    phrase(i) = trim(phrase(i))
		  next i
		  if phrase.Ubound = 3 then
		    hintOut = val(phrase(0))
		    tableNameOut = phrase(1)
		    forOut = phrase(2)
		    graphNameOut = phrase(3)
		    locFirstDashes = instr(graphNameOut, "--")
		    locSecondDashes = instr(locFirstDashes + 2, graphNameOut, "--")
		    IF locFirstDashes > 1 then
		      headOneOut = trim(graphNameOut.Left(locFirstDashes - 1))
		      if locSecondDashes > 1 then
		        headTwoOut = trim(graphNameOut.Mid(locFirstDashes + 2,locSecondDashes - (locFirstDashes +3)))
		        kindGraphOut = trim(graphNameOut.Mid(locSecondDashes+2))
		      else
		        kindGraphOut = trim(graphNameOut.Mid(locFirstDashes+2))
		        headTwoOut = ""
		      end if
		    else
		      headOneOut = ""
		      headTwoOut = ""
		      kindGraphOut = ""
		    end if
		  else
		    hintOut = -1
		    tableNameOut = ""
		    forOut = ""
		    graphNameOut = ""
		    headOneOut = ""
		    headTwoOut = ""
		    kindGraphOut = ""
		  end if
		  i = 1
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub DrawChart()
		  'currentGraphCode acts as a parameter to this routine even though it is a MainModule property.
		  dim numFiles as integer
		  dim iFile as Integer = 0
		  dim curFile as FolderItem
		  dim fileNameLabels(-1) as String
		  dim selHint as integer = 0
		  dim selTableName, selFor,selGraphName, selHeadOne, selHeadTwo, selKindGraph as String
		  dim jHead as integer = 0
		  dim jMonth as integer = 0
		  dim tableReference(-1) as Integer
		  dim tableColumns(-1) as Integer
		  dim tableRows(-1) as integer
		  dim dataForGraph(-1,-1)  as Single 'files , legend
		  dim legendItems(-1) as String
		  dim kLegend as integer = 0
		  dim dataSeries(-1) as double
		  dim graphColors(39) as integer
		  dim randomNumber as new Random
		  'chart related items
		  dim c as new CDXYChartMBS(plotCanvas.Width,plotCanvas.Height)
		  dim layer as CDBarLayerMBS
		  dim lineLayer as CDLineLayerMBS
		  dim lb as CDLegendBoxMBS
		  dim monthNames(11) as string
		  dim showNoGraphAvail as Boolean
		  'dim noGraphPic as Picture
		  dim curFileString as String
		  dim curFileStringLen as Integer
		  
		  showNoGraphAvail = false
		  graphColors(0) = &h8080FF
		  graphColors(1) = &h800080
		  graphColors(2) = &hFFFFB9
		  graphColors(3) = &h80FFFF
		  graphColors(4) = &h800040
		  graphColors(5) = &hFF8040
		  graphColors(6) = &h0000FF
		  graphColors(7) = &hE4F2F3
		  graphColors(8) = &h400080
		  graphColors(9) = &hFF00FF
		  graphColors(10) = &hFFFF80
		  graphColors(11) = &hF2F2FD
		  graphColors(12) = &h800080
		  graphColors(13) = &h800000
		  graphColors(14) = &h008000
		  graphColors(15) = &h0000FF
		  graphColors(16) = &hFF0000
		  graphColors(17) = &hCEFF9D
		  graphColors(18) = &hDFDFFF
		  graphColors(19) = &hECEC00
		  'repeat colors
		  graphColors(20) = &h8080FF
		  graphColors(21) = &h800080
		  graphColors(22) = &hFFFFB9
		  graphColors(23) = &h80FFFF
		  graphColors(24) = &h800040
		  graphColors(25) = &hFF8040
		  graphColors(26) = &h0000FF
		  graphColors(27) = &hE4F2F3
		  graphColors(28) = &h400080
		  graphColors(29) = &hFF00FF
		  graphColors(30) = &hFFFF80
		  graphColors(31) = &hF2F2FD
		  graphColors(32) = &h800080
		  graphColors(33) = &h800000
		  graphColors(34) = &h008000
		  graphColors(35) = &h0000FF
		  graphColors(36) = &hFF0000
		  graphColors(37) = &hCEFF9D
		  graphColors(38) = &hDFDFFF
		  graphColors(39) = &hECEC00
		  
		  monthNames(0) = "Jan"
		  monthNames(1) = "Feb"
		  monthNames(2) = "Mar"
		  monthNames(3) = "Apr"
		  monthNames(4) = "May"
		  monthNames(5) = "Jun"
		  monthNames(6) = "Jul"
		  monthNames(7) = "Aug"
		  monthNames(8) = "Sep"
		  monthNames(9) = "Oct"
		  monthNames(10) = "Nov"
		  monthNames(11) = "Dec"
		  
		  
		  numFiles = ActiveFileList.Ubound
		  'Decode the selected graph
		  'codeIn as string, hintOut as integer, tableNameOut as string, forOut as string, graphNameOut as string, headOneOut as string, headTwoOut as string, kindGraphOut as String
		  CALL DecodeGraphName(currentGraphCode,selHint, selTableName, selFor,selGraphName, selHeadOne, selHeadTwo, selKindGraph)
		  if selHint >= 0 then
		    'make the labels the names of each file
		    redim fileNameLabels(numFiles)
		    for iFile = 0 to numFiles
		      curFile = new FolderItem(ActiveFileList(iFile))
		      curFileString = curFile.DisplayName
		      curFileStringLen =curfilestring.Len
		      if curFileStringLen > (plotCanvas.width-350) / ((numFiles + 1)*5)  then
		        if curFileStringLen mod 2 = 1 then curFileStringLen=curFileStringLen + 1
		        curFileString = curFileString.Left(curFileStringLen/2) + EndOfLine.UNIX + curFileString.Mid(1+curFileStringLen/2)
		        'curFileString = curFileString + EndOfLine.UNIX + "This is the second line"
		      end if
		      fileNameLabels(iFile) = curFileString
		    next iFile
		    'get the list of table references for each file
		    tableReference = GetSpecificTableReference(selHint,selFor)
		    'all types of graphs are handled similarly except for monthly
		    if selKindGraph = "Monthly Line" Then
		      call c.setPlotArea(30,20, plotCanvas.Width-60, plotCanvas.Height-50)
		      call c.addLegend(150, 100)
		      call c.addTitle(selTableName + " -- " + selGraphName, "times.ttf", 12)
		      graphColors(2) = &h808080  ' 'don't use yellow line since it is too faint
		      call c.xAxis.setLabels(monthNames)
		      lineLayer = c.addLineLayer
		      tableColumns = GetSpecificTableColumn(tableReference, selHeadOne)
		      dataForGraph = GetDataForMonthlyGraph(tableReference,tableColumns)
		      redim dataSeries(11) 'one for each month
		      for iFile = 0 to numFiles
		        for jMonth = 0 to 11
		          dataSeries(jMonth) = dataForGraph(iFile,jMonth)
		        next jMonth
		        call lineLayer.addDataSet(dataSeries, graphColors(iFile), fileNameLabels(iFile))
		      next iFile
		      pictOfChart=c.makeChartPicture
		      plotcanvas.Backdrop  = pictOfChart
		    else 'all graphs except monthly
		      'legend on the side of graph never really worked because the size of the legend varied
		      call c.setPlotArea(50,40, plotCanvas.Width-250, plotCanvas.Height-80)
		      'set up legend box
		      lb = c.addLegend(plotCanvas.Width-10, 100)
		      call lb.setAlignment(9) ' topright
		      call lb.setBackground(&hFFFFFF)
		      'now the title and axes
		      call c.addTitle(selTableName + " -- " + selGraphName, "times.ttf", 12)
		      call c.xAxis.setLabels(fileNameLabels)
		      'call c.xAxis.setLabelStyle("",8,&hffff0002,5)
		      call c.yaxis.SetAutoScale(0.1,0.1,1)
		      select case selKindGraph
		      case "Stacked Bar"
		        layer = c.addBarLayer(c.kStack, 0)
		        'try first that it is a column  heading
		        tableColumns = GetSpecificTableColumn(tableReference,selHeadOne)
		        if tableColumns.Ubound = -1 then 'it must be a row heading not a column heading
		          tableRows = GetSpecificTableRow(tableReference,selHeadOne)
		        end if
		      case "100% Stacked Bar"
		        layer = c.addBarLayer(c.kPercentage, 0)
		        'try first that it is a column  heading
		        tableColumns = GetSpecificTableColumn(tableReference,selHeadOne)
		        if tableColumns.Ubound = -1 then 'it must be a row heading not a column heading
		          tableRows = GetSpecificTableRow(tableReference,selHeadOne)
		        end if
		      case "Side-by-side Bar"
		        layer = c.addBarLayer(c.kSide, 0)
		        'try first that it is a column  heading
		        tableColumns = GetSpecificTableColumn(tableReference,selHeadOne)
		        if tableColumns.Ubound = -1 then 'it must be a row heading not a column heading
		          tableRows = GetSpecificTableRow(tableReference,selHeadOne)
		        end if
		      case "Simple Bar"
		        layer = c.addBarLayer(c.kStack, 0)
		        tableColumns = GetSpecificTableColumn(tableReference,selHeadTwo)
		        tableRows = GetSpecificTableRow(tableReference,selHeadOne)
		      else
		        showNoGraphAvail = true
		        'layer = c.addBarLayer(c.kStack, 0)
		      end select
		      if tableColumns.Ubound >=0 or tablerows.Ubound >= 0 then
		        legendItems = GetLegendForGraph(tableReference,tableRows,tableColumns)
		        dataForGraph = GetDataForGraph(tableReference,tableRows,tableColumns,legendItems)
		        redim dataSeries(numFiles)
		        for kLegend = 0 to legendItems.Ubound
		          for iFile = 0 to numFiles
		            dataSeries(iFile) = dataForGraph(iFile,kLegend)
		            'if dataForGraph(iFile,kLegend)>1000 then
		            'dataSeries(iFile) = round(dataForGraph(iFile,kLegend))
		            'else
		            'dataSeries(iFile) = dataForGraph(iFile,kLegend)
		            'end if
		          next iFile
		          call layer.addDataSet(dataSeries, graphColors(kLegend mod graphColors.ubound), legendItems(kLegend))
		          call layer.setDataLabelStyle
		          call layer.setDataLabelFormat("{value|2,.}")
		          if selKindGraph = "Stacked Bar" then
		            call layer.setAggregateLabelStyle
		            call layer.setAggregateLabelFormat("{value|2,.}")
		          end if
		        next kLegend
		        pictOfChart=c.makeChartPicture
		        plotcanvas.Backdrop  = pictOfChart
		      else
		        showNoGraphAvail = true
		      end if
		    end if
		  else
		    showNoGraphAvail = true
		  end if
		  if showNoGraphAvail then
		    'plotCanvas.backdrop =
		    dim noGraphPic as  new  Picture(plotCanvas.Width, plotCanvas.Height)
		    noGraphPic.Graphics.TextFont="Helvetica"
		    noGraphPic.Graphics.TextSize=18
		    noGraphPic.Graphics.ForeColor=&c000000
		    noGraphPic.Graphics.DrawString "No graph available",100,100
		    'plotCanvas.DrawPicture(noGraphPic,0,0)
		  end if
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetDataForGraph(tableRefIn() as integer, tabRowIn() as integer, tabColIn() as integer, LegendIn() as string) As single(,)
		  dim DataToGraphOut(-1,-1) as single 'files, legend items
		  dim numFiles as integer = 0
		  dim numLegend as integer = 0
		  dim iFile as Integer = 0
		  dim jHead as Integer = 0
		  dim curTableRef as Integer = 0
		  dim kLegend as Integer = 0
		  dim curCol as Integer = 0
		  dim curRow as Integer = 0
		  dim curValue as single = 0
		  dim curColHead as String
		  dim curRowHead as String
		  dim legendSplit as Integer
		  
		  numFiles = ActiveFileList.Ubound
		  numLegend = LegendIn.Ubound
		  redim DataToGraphOut(numFiles,numLegend)
		  'make sure all data items are zero
		  for iFile = 0 to numFiles
		    for kLegend = 0 to numLegend
		      DataToGraphOut(iFile,kLegend) =  0.0
		    next kLegend
		  next iFile
		  if numFiles = tableRefIn.Ubound and numLegend >= 0 then
		    if tabRowIn.Ubound>=0 and tabColIn.Ubound >= 0 then 'both rows and colums are defined so a single legend item should be created
		      for iFile = 0 to numFiles
		        curTableRef = tableRefIn(iFile)
		        if curTableRef > 0 and curTableRef <= TableResults.Ubound then
		          legendSplit = legendin(0).instr(" -- ")
		          if legendSplit > 0 then
		            curColHead = LegendIn(0).Left(legendSplit-1)
		            curCol = 0
		            for jHead = 1 to TableResults(curTableRef).NumberOfColumns
		              if curColHead = Nm(jHead + TableResults(curTableRef).FirstColHeadNameIndex -1) then
		                curCol = jHead
		                exit For jHead
		              end if
		            next jHead
		            curRowHead = LegendIn(0).Mid(legendSplit + 4)
		            curRow = 0
		            for jHead = 1 to TableResults(curTableRef).NumberOfRows
		              if curRowHead = Nm(jHead + TableResults(curTableRef).FirstRowHeadNameIndex - 1) then
		                curRow = jHead
		                exit for jHead
		              end if
		            next jHead
		            if curCol>0 and curRow> 0 then
		              curValue = v(TableResults(curTableRef).FirstValueIndex + (curCol-1) + (curRow-1)*TableResults(curTableRef).NumberOfColumns)
		            else
		              curValue = 0
		            end if
		            DataToGraphOut(iFile,0) = curValue 'only one item in the legend array so it is the first in DataToGraphOut
		          end if
		        end if
		      next ifile
		    elseif tabRowIn.Ubound >=0 then 'a row is defined but not a column so go through all of the columns for legend items
		      for iFile = 0 to numFiles
		        curTableRef = tableRefIn(iFile)
		        if curTableRef > 0 and curTableRef <= TableResults.Ubound then
		          'search for column headers that match legend
		          for kLegend = 0 to numLegend
		            curCol = 0
		            for jHead = 1 to TableResults(curTableRef).NumberOfColumns
		              if LegendIn(kLegend) = Nm(jHead + TableResults(curTableRef).FirstColHeadNameIndex -1) then
		                curCol = jHead
		                exit For jHead
		              end if
		            next jHead
		            if curCol>0 then
		              curRow = tabRowIn(iFile)
		              curValue = v(TableResults(curTableRef).FirstValueIndex + (curCol-1) +(curRow-1)*TableResults(curTableRef).NumberOfColumns)
		            else
		              curValue = 0
		            end if
		            DataToGraphOut(iFile,kLegend) = curValue
		          next kLegend
		        end if
		      next ifile
		    elseif tabColIn.Ubound >= 0 then
		      for iFile = 0 to numFiles
		        curTableRef = tableRefIn(iFile)
		        if curTableRef > 0 and curTableRef <= TableResults.Ubound then
		          for kLegend = 0 to numLegend
		            curRow= 0
		            for jHead = 1 to TableResults(curTableRef).NumberOfRows
		              if LegendIn(kLegend) = Nm(jHead + TableResults(curTableRef).FirstRowHeadNameIndex - 1) then
		                curRow = jHead
		                exit for jHead
		              end if
		            next jHead
		            if curRow>0 then
		              curCol = tabColIn(iFile)
		              curValue = v(TableResults(curTableRef).FirstValueIndex + (curCol-1) +(curRow-1)*TableResults(curTableRef).NumberOfColumns)
		            else
		              curValue = 0
		            end if
		            DataToGraphOut(iFile,kLegend) = curValue
		          next kLegend
		        end if
		      next ifile
		    else
		      MsgBox "Both rows and columns are not defined for getting data for the graph"
		    end if
		  else
		    MsgBox "Number of legend items and number of files not correct"
		  end if
		  
		  
		  return DataToGraphOut
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetDataForMonthlyGraph(tableRefIn() as integer, tabColIn() as integer) As single(,)
		  dim DataToGraphOut(-1,-1) as single 'files, legend items
		  dim numFiles as integer = 0
		  dim numLegend as integer = 0
		  dim iFile as Integer = 0
		  dim jHead as Integer = 0
		  dim curTableRef as Integer = 0
		  dim kMonth as Integer = 0
		  dim curCol as Integer = 0
		  dim curRow as Integer = 0
		  dim curValue as single = 0
		  dim curColHead as String
		  dim curRowHead as String
		  dim legendSplit as Integer
		  
		  numFiles = ActiveFileList.Ubound
		  redim DataToGraphOut(numFiles,11)
		  'make sure all data items are zero
		  for iFile = 0 to numFiles
		    for kMonth = 0 to numLegend
		      DataToGraphOut(iFile,kMonth) =  0.0
		    next kMonth
		  next iFile
		  if tabColIn.Ubound > 0 then
		    for iFile = 0 to numFiles
		      curTableRef = tableRefIn(iFile)
		      if curTableRef > 0 and curTableRef <= TableResults.Ubound then
		        if tabColIn.Ubound > 0 then
		          curCol = tabColIn(iFile)
		          for kMonth = 1 to 12
		            curRow = kMonth
		            curValue = v(TableResults(curTableRef).FirstValueIndex + (curCol-1) +(curRow-1)*TableResults(curTableRef).NumberOfColumns)
		            DataToGraphOut(iFile,kMonth-1) = curValue
		          next kMonth
		        end if
		      end if
		    next ifile
		  else
		    MsgBox "Monthly graph data expected references to columns."
		  end if
		  return DataToGraphOut
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetLegendForGraph(tableRefIn() as integer, tabRowIn() as integer, tabColIn() as integer) As string()
		  Dim LegendOut(-1) as string
		  dim curLegendItem as String
		  
		  dim numFiles as integer = 0
		  dim iFile as Integer = 0
		  dim jHead as Integer = 0
		  dim curTableRef as integer = 0
		  dim curHint as integer
		  dim exclLeftCols as Integer
		  dim exclRightCols as Integer
		  dim exclBotRow as Integer
		  dim exclTopRow as Integer
		  dim nmIndexCol as Integer
		  dim nmIndexRow as Integer
		  
		  numFiles = ActiveFileList.Ubound
		  if numFiles = tableRefIn.Ubound then
		    if ubound(tabRowIn) >=0 and ubound(tabColIn) >= 0 then 'both rows and colums are defined so a single legend item should be created
		      for iFile = 0 to numFiles
		        curTableRef = tableRefIn(iFile)
		        if curTableRef >= 0 and curTableRef <= TableResults.Ubound then
		          nmIndexCol = tabColin(iFile) + TableResults(curTableRef).FirstColHeadNameIndex - 1
		          nmIndexRow = tabRowIn(iFile) + TableResults(curTableRef).FirstRowHeadNameIndex - 1
		          if nmIndexCol<= ubound(nm) and nmIndexCol>=0 and nmIndexRow<=ubound(nm) and nmIndexRow>=0 then
		            curLegendItem = Nm(nmIndexCol) + " -- " + Nm(nmIndexRow)
		            if curLegendItem <> "&nbsp;" then
		              if LegendOut.IndexOf(curLegendItem) = - 1 then
		                LegendOut.Append curLegendItem
		              end if
		            end if
		          end if
		        end if
		      next ifile
		      if LegendOut.Ubound > 0 then 'should only find one unique item
		        MsgBox "Error: should have found only one unique item"
		      elseif LegendOut.ubound = -1 then 'should have one unique item
		        msgbox "Error: no legend item created for case with both rows and columns defined."
		      end if
		    elseif ubound(tabRowIn) >=0 then 'a row is defined but not a column so go through all of the columns for legend items
		      for iFile = 0 to numFiles
		        curTableRef = tableRefIn(iFile)
		        if curTableRef >= 0 and curTableRef <= TableResults.Ubound then
		          curHint = TableResults(curTableRef).GraphHintIndex
		          if curHint>=0 and curHint<GraphHints.Ubound then
		            exclLeftCols = GraphHints(curHint).numLeftColumnsToExclude
		            exclRightCols = GraphHints(curHint).numRightColumnsToExclude
		          else
		            exclLeftCols = 0
		            exclRightCols = 0
		          end if
		          for jHead = 1 + exclLeftCols to TableResults(curTableRef).NumberOfColumns - exclRightCols
		            curLegendItem = Nm(jHead + TableResults(curTableRef).FirstColHeadNameIndex - 1)
		            if curLegendItem <> "&nbsp;" then
		              if LegendOut.IndexOf(curLegendItem) = - 1 then
		                LegendOut.Append  curLegendItem
		              end if
		            end if
		          next jHead
		        end if
		      next ifile
		    elseif ubound(tabColIn) >= 0 then
		      for iFile = 0 to numFiles
		        curTableRef = tableRefIn(iFile)
		        if curTableRef >= 0 and curTableRef <= TableResults.Ubound then
		          curHint = TableResults(curTableRef).GraphHintIndex
		          if curHint>=0 and curHint<GraphHints.Ubound then
		            exclBotRow = GraphHints(curHint).numBottomRowsToExclude
		            exclTopRow = GraphHints(curHint).numTopRowsToExclude
		          else
		            exclBotRow = 0
		            exclTopRow = 0
		          end if
		          for jHead = 1 + exclTopRow to TableResults(curTableRef).NumberOfRows - exclBotRow
		            curLegendItem = Nm(jHead + TableResults(curTableRef).FirstRowHeadNameIndex - 1)
		            if curLegendItem <> "&nbsp;" then
		              if LegendOut.IndexOf(curLegendItem) = - 1 then
		                LegendOut.Append  curLegendItem
		              end if
		            end if
		          next jHead
		        end if
		      next ifile
		    end if
		  else
		    MsgBox "Both rows and columns are not defined for getting the legend items"
		  end if
		  return LegendOut
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub GetSpecificGraphNames(inHintIndex as integer, listOut() as string, graphCodeOut() as String)
		  ' these arrays hold all row and column headings
		  dim rowHeadings(-1) as string
		  dim colHeadings(-1) as string
		  dim lastRowHeading as integer = 0
		  dim lastColHeading as integer = 0
		  'these arrays hold only the row and column headings that have not been excluded
		  dim rowNonExclHeadings(-1) as string
		  dim colNonExclHeadings(-1) as string
		  dim lastRowNonExclHeading as integer = 0
		  dim lastColNonExclHeading as integer = 0
		  'other variables
		  dim numFiles as integer = 0
		  dim curHeading as String
		  dim iTable as Integer = 0
		  dim jFile as integer = 0
		  dim kRow as integer = 0
		  dim kRowOffset as integer = 0
		  dim lCol as Integer = 0
		  dim lColOffset as integer = 0
		  
		  numFiles = ActiveFileList.Ubound
		  for jFile = 0 to numFiles
		    if HintToTableStart(inHintIndex ,jFile) >= 0 and HintToTableEnd(inHintIndex,jFile) >= 0 then
		      for iTable = HintToTableStart(inHintIndex,jFile) to HintToTableEnd(inHintIndex,jFile)
		        if iTable>=0 and iTable<= TableResults.Ubound then
		          if TableResults(iTable).GraphHintIndex = inHintIndex then
		            
		            'gather the list of all row headings for the selected table
		            for kRow = 1 to TableResults(iTable).NumberOfRows
		              kRowOffset = kRow + TableResults(iTable).FirstRowHeadNameIndex - 1
		              curHeading = Nm(kRowOffset)
		              if curHeading <> "&nbsp;" then
		                if rowHeadings.IndexOf(curHeading) = -1 then
		                  rowHeadings.Append curHeading
		                end if
		              end if
		            next kRow
		            
		            'gather the list of non-excluded row headings for the selected table
		            for kRow = GraphHints(inHintIndex).numTopRowsToExclude + 1 to TableResults(iTable).NumberOfRows -  GraphHints(inHintIndex).numBottomRowsToExclude
		              kRowOffset = kRow + TableResults(iTable).FirstRowHeadNameIndex - 1
		              curHeading = Nm(kRowOffset)
		              if curHeading <> "&nbsp;" then
		                if rowNonExclHeadings.IndexOf(curHeading) = -1 then
		                  rowNonExclHeadings.Append curHeading
		                end if
		              end if
		            next kRow
		            
		            'gather the list of all column headings for the selected table
		            for lCol = 1 to TableResults(iTable).NumberOfColumns
		              lColOffset = lCol + TableResults(iTable).FirstColHeadNameIndex  - 1
		              curHeading = Nm(lColOffset)
		              if curHeading <> "&nbsp;" then
		                if colHeadings.IndexOf(curHeading) = -1 then
		                  colHeadings.Append curHeading
		                end if
		              end if
		            next lCol
		            
		            'gather the list of all non-excluded column headings for the selected table
		            for lCol = GraphHints(inHintIndex).numLeftColumnsToExclude + 1 to TableResults(iTable).NumberOfColumns - GraphHints(inHintIndex).numRightColumnsToExclude
		              lColOffset = lCol + TableResults(iTable).FirstColHeadNameIndex  - 1
		              curHeading = Nm(lColOffset)
		              if curHeading <> "&nbsp;" then
		                if colNonExclHeadings.IndexOf(curHeading) = -1 then
		                  colNonExclHeadings.Append curHeading
		                end if
		              end if
		            next lCol
		          end if
		        end if
		      next iTable
		    end if
		  next jFile
		  
		  lastRowHeading = rowHeadings.Ubound
		  lastColHeading = colHeadings.Ubound
		  lastRowNonExclHeading = rowNonExclHeadings.Ubound
		  lastColNonExclHeading = colNonExclHeadings.Ubound
		  
		  'Stacked bar of values from each column
		  if GraphHints(inHintIndex).isStackedBarForEachColumn then
		    for lCol = 0 to lastColNonExclHeading
		      listOut.Append colNonExclHeadings(lCol) + " -- Stacked Bar"
		    next lCol
		  end if
		  
		  'Stacked bar of values from each row
		  if GraphHints(inHintIndex).isStackedBarForEachRow then
		    for kRow = 0 to lastRowNonExclHeading
		      listOut.Append rowNonExclHeadings(kRow) + " -- Stacked Bar"
		    next kRow
		  end if
		  
		  '100% stacked bar of values from each column
		  if GraphHints(inHintIndex).is100StackedBarForEachColumn then
		    for lCol = 0 to lastColNonExclHeading
		      listOut.Append colNonExclHeadings(lCol) + " -- 100% Stacked Bar"
		    next lCol
		  end if
		  
		  '100% stacked bar of values from each row
		  if GraphHints(inHintIndex).is100StackedBarForEachRow then
		    for kRow = 0 to lastRowNonExclHeading
		      listOut.Append rowNonExclHeadings(kRow) + " -- 100% Stacked Bar"
		    next kRow
		  end if
		  
		  'Side-by-side bar of values from each column
		  if GraphHints(inHintIndex).isSideBySideBarForEachColumn then
		    for lCol = 0 to lastColNonExclHeading
		      listOut.Append colNonExclHeadings(lCol) + " -- Side-by-side Bar"
		    next lCol
		  end if
		  
		  'Side-by-side bar of values from each row
		  if GraphHints(inHintIndex).isSideBySideBarForEachRow then
		    for kRow = 0 to lastRowNonExclHeading
		      listOut.Append rowNonExclHeadings(kRow) + " -- Side-by-side Bar"
		    next kRow
		  end if
		  
		  'Side-by-side bar of each value in total section across instance of report in file
		  
		  'Monthly line graph for values from each column
		  if GraphHints(inHintIndex).isMonthlyLineForEachColumn then
		    for lCol = 0 to lastColNonExclHeading
		      listOut.Append colNonExclHeadings(lCol) + " -- Monthly Line"
		    next lCol
		  end if
		  
		  
		  'Show simple bar graphs for every individual value in table
		  if GraphHints(inHintIndex).isBarForEveryValue then
		    for kRow = 0 to lastRowHeading
		      for lCol = 0 to lastColHeading
		        listOut.Append rowHeadings(kRow) + " -- " + colHeadings(lCol) + " -- Simple Bar"
		      next lCol
		    next kRow
		  end if
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetSpecificTableColumn(tableRefIn() as integer, HeadingIn as String) As Integer()
		  dim columnOut(-1) as integer
		  dim numFiles as integer = 0
		  dim iFile as Integer = 0
		  dim jHead as integer = 0
		  dim curTableRef as integer = 0
		  dim foundAny as Boolean
		  numFiles = ActiveFileList.Ubound
		  
		  foundAny = False
		  redim columnOut(numFiles)
		  if numFiles = tableRefIn.Ubound then
		    for iFile = 0 to numFiles
		      columnOut(iFile) = 0
		      curTableRef = tableRefIn(iFile)
		      if curTableRef > 0 and curTableRef <= TableResults.Ubound then
		        for jHead = 1 to TableResults(curTableRef).NumberOfColumns
		          if HeadingIn  = Nm(jHead + TableResults(curTableRef).FirstColHeadNameIndex - 1) then
		            columnOut(iFile) = jHead
		            foundAny = True
		            exit for jHead
		          end if
		        next jHead
		      end if
		    next iFile
		  end if
		  if not foundAny then
		    redim columnOut(-1)
		  end if
		  return columnOut
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetSpecificTableReference(hintIn as Integer, forIn as string) As Integer()
		  ' given the hint and for string return a table reference for each active file
		  dim numFiles as integer = 0
		  dim tabRef(-1) as integer
		  dim iFile as Integer = 0
		  dim kTable as integer = 0
		  numFiles = ActiveFileList.Ubound
		  
		  redim tabRef(numFiles)
		  if hintIn >= 0 then
		    for iFile = 0 to numFiles
		      if HintToTableStart(hintIn,iFile) >= 0 and HintToTableEnd(hintIn,iFile) >= 0 then
		        if forIn = "NONE" then
		          tabRef(iFile) = HintToTableStart(hintIn ,iFile) 'if only one table just use first one
		        else
		          for kTable = HintToTableStart(hintIn ,iFile) to HintToTableEnd(hintIn ,iFile)
		            if kTable>=0 and kTable<=TableResults.Ubound then
		              if TableResults(kTable).GraphHintIndex = hintIn  then
		                if forIn  =  Nm(TableResults(kTable).ForNameIndex) then
		                  tabRef(iFile) = kTable
		                  exit for
		                end if
		              end if
		            end if
		          next kTable
		        end if
		      end if
		    next iFile
		  end if
		  return tabRef
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetSpecificTableRow(tableRefIn() as integer, HeadingIn as String) As Integer()
		  dim rowOut(-1) as integer
		  dim numFiles as integer = 0
		  dim iFile as Integer = 0
		  dim jHead as integer = 0
		  dim curTableRef as integer = 0
		  numFiles = ActiveFileList.Ubound
		  
		  redim rowOut(numFiles)
		  if numFiles = tableRefIn.Ubound then
		    for iFile = 0 to numFiles
		      rowOut(iFile) = 0
		      curTableRef = tableRefIn(iFile)
		      if curTableRef > 0 and curTableRef <= TableResults.Ubound then
		        for jHead = 1 to TableResults(curTableRef).NumberOfRows
		          if HeadingIn  = Nm(jHead + TableResults(curTableRef).FirstRowHeadNameIndex - 1) then
		            rowOut(iFile) = jHead
		            exit for jHead
		          end if
		        next jHead
		      end if
		    next iFile
		  end if
		  return rowOut
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub PopulateList()
		  dim tableNode as TreeViewNode
		  dim forNode as TreeViewNode
		  dim graphNode as TreeViewNode
		  dim listItem as string = ""
		  dim iHint as integer = 0
		  dim counter as integer = 10000
		  dim jFile as integer = 0
		  dim kTable as Integer = 0
		  dim mSpGrph as integer = 0
		  dim nFor as Integer = 0
		  dim numFiles as integer = 0
		  dim curFor as String
		  dim uniqueFors(-1) as string
		  dim specificGraphNames(-1) as String
		  dim specificGraphCodes(-1) as String
		  dim firstRootNode as TreeViewNode
		  
		  TreeViewGraphs.LockDrawing = true
		  TreeViewGraphs.RemoveAllNodes
		  numFiles = ActiveFileList.Ubound
		  for iHint = 0 to ubound(ghCombinedName)
		    counter = counter + 1
		    tableNode = new TreeViewNode(ghCombinedName(iHint))
		    redim specificGraphNames(-1)
		    call getSpecificGraphNames(iHint,specificGraphNames ,specificGraphCodes)
		    'find the "fors" unless they are for the entire building
		    redim uniqueFors(-1)
		    for jFile = 0 to numFiles
		      if HintToTableStart(iHint,jFile) >= 0 and HintToTableEnd(iHint,jFile) >= 0 then
		        for kTable = HintToTableStart(iHint,jFile) to HintToTableEnd(iHint,jFile)
		          if kTable>=0 and kTable<= TableResults.Ubound then
		            if TableResults(kTable).GraphHintIndex = iHint then
		              curFor = Nm(TableResults(kTable).ForNameIndex)
		              if curFor <> "Entire Facility" and curFor <> "Meter" then
		                if uniqueFors.IndexOf(curFor) = -1 then
		                  uniqueFors.Append curFor
		                end if
		              end if
		            end if
		          end if
		        next kTable
		      end if
		    next jFile
		    if uniqueFors.Ubound > 0 then
		      for nFor = 0 to uniqueFors.Ubound
		        forNode = new TreeViewNode(uniqueFors(nFor))
		        for mSpGrph = 0 to specificGraphNames.Ubound
		          graphNode = new TreeViewNode(specificGraphNames(mSpGrph))
		          graphNode.ItemData = str(iHint) + " | " + ghCombinedName(iHint) + " | " + uniqueFors(nFor) + " | " + specificGraphNames(mSpGrph)
		          forNode.AppendNode(graphNode)
		        next mSpGrph
		        tableNode.AppendNode(forNode)
		      next nFor
		    else
		      for mSpGrph = 0 to specificGraphNames.Ubound
		        graphNode = new TreeViewNode(specificGraphNames(mSpGrph))
		        graphNode.ItemData = str(iHint) + " | " + ghCombinedName(iHint) + " | NONE | " + specificGraphNames(mSpGrph)
		        tableNode.AppendNode(graphNode)
		      next mSpGrph
		    end if
		    TreeViewGraphs.AppendNode(tableNode)
		  next iHint
		  TreeViewGraphs.LockDrawing = false
		  firstRootNode = TreeViewGraphs.RootNodes(0)
		  firstRootNode.SetExpanded(true,true)
		  if firstRootNode.NodeCount>1 then
		    TreeViewGraphs.SelectedIndex = 1
		  end  if
		  
		  
		End Sub
	#tag EndMethod


#tag EndWindowCode

#tag Events tbMain1
	#tag Event
		Sub Action(item As ToolItem)
		  dim clip as Clipboard
		  
		  select case Item.Name
		  case "tbiManageFiles"
		    ManageFilesWindow.ShowModal
		    me.MouseCursor = system.Cursors.Wait
		    Call ClearMainArrays
		    Call LoadAllFiles
		    Call PopulateList
		    me.MouseCursor = system.Cursors.StandardPointer
		    CurrentGraphCode = " | | | "
		    Call DrawChart
		  case "tbiShowGraph"
		    'msgbox "Show Graph"
		    if ToolButton(item).Pushed then
		      PlotCanvas.Visible = true
		      lstResults.Visible = false
		    else
		      PlotCanvas.Visible = false
		      lstResults.Visible = true
		    end if
		  case "tbiCopy"
		    clip = new Clipboard
		    'clip.SetText ""
		    clip.Picture = pictOfChart
		    clip.Close
		  case "tbiNext"
		    MsgBox "Next not yet implemented"
		  case "tbiPrevious"
		    MsgBox "Previous not yet implemented"
		  case "tbiFavoritesOnly"
		    MsgBox "Favorites Only not yet implemented"
		  case "tbiAddFavorite"
		    MsgBox "Add Favorite not yet implemented"
		  case "tbiRemoveFavorite"
		    MsgBox "Remove Favorite not yet implemented"
		  case "tbiAbout"
		    AboutWindow.ShowModal
		    'MsgBox "EP-Compare Version 0.2" + EndOfLine + _
		    '"Copyright  (c) 209-2010 GARD Analytics, Inc. " + EndOfLine + _
		    '"All rights reserved. See Notice in EP-Launch for all terms."
		  end select
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events TreeViewGraphs
	#tag Event
		Sub SelectionChanged()
		  Dim node as TreeViewNode
		  
		  if me.SelectedIndex >= 0 then
		    node = me.SelectedNode()
		    currentGraphCode = node.ItemData
		    Call DrawChart
		  end if
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag ViewBehavior
	#tag ViewProperty
		Name="Name"
		Visible=true
		Group="ID"
		Type="String"
		EditorType="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Interfaces"
		Visible=true
		Group="ID"
		Type="String"
		EditorType="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Super"
		Visible=true
		Group="ID"
		Type="String"
		EditorType="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Width"
		Visible=true
		Group="Size"
		InitialValue="600"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Height"
		Visible=true
		Group="Size"
		InitialValue="400"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinWidth"
		Visible=true
		Group="Size"
		InitialValue="64"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinHeight"
		Visible=true
		Group="Size"
		InitialValue="64"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaxWidth"
		Visible=true
		Group="Size"
		InitialValue="32000"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaxHeight"
		Visible=true
		Group="Size"
		InitialValue="32000"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Frame"
		Visible=true
		Group="Frame"
		InitialValue="0"
		Type="Integer"
		EditorType="Enum"
		#tag EnumValues
			"0 - Document"
			"1 - Movable Modal"
			"2 - Modal Dialog"
			"3 - Floating Window"
			"4 - Plain Box"
			"5 - Shadowed Box"
			"6 - Rounded Window"
			"7 - Global Floating Window"
			"8 - Sheet Window"
			"9 - Metal Window"
			"11 - Modeless Dialog"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Title"
		Visible=true
		Group="Frame"
		InitialValue="Untitled"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="CloseButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Resizeable"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaximizeButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinimizeButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="FullScreenButton"
		Visible=true
		Group="Frame"
		InitialValue="False"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Composite"
		Group="OS X (Carbon)"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MacProcID"
		Group="OS X (Carbon)"
		InitialValue="0"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="FullScreen"
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="ImplicitInstance"
		Visible=true
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LiveResize"
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Placement"
		Visible=true
		Group="Behavior"
		InitialValue="0"
		Type="Integer"
		EditorType="Enum"
		#tag EnumValues
			"0 - Default"
			"1 - Parent Window"
			"2 - Main Screen"
			"3 - Parent Window Screen"
			"4 - Stagger"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Visible"
		Visible=true
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasBackColor"
		Visible=true
		Group="Background"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="BackColor"
		Visible=true
		Group="Background"
		InitialValue="&hFFFFFF"
		Type="Color"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Backdrop"
		Visible=true
		Group="Background"
		Type="Picture"
		EditorType="Picture"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBar"
		Visible=true
		Group="Menus"
		Type="MenuBar"
		EditorType="MenuBar"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBarVisible"
		Visible=true
		Group="Deprecated"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
#tag EndViewBehavior
