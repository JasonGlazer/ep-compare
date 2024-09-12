#tag Module
Protected Module MainModule
	#tag Method, Flags = &h0
		Function AddName(inStringName as String) As integer
		  CountNm = CountNm + 1
		  If CountNm>Nm.Ubound then
		    redim Nm(CountNm + 100)
		  end if
		  Nm(CountNm) = inStringName
		  return CountNm
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function AddValueFromString(inValueAsString as String) As Integer
		  CountV = CountV + 1
		  if CountV >V.Ubound then
		    redim V(CountV + 100)
		  end if
		  V(CountV) = val(inValueAsString)
		  return CountV
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ClearMainArrays()
		  CountNm = 0
		  CountV = 0
		  CountTableResults = 0
		  Redim TableResults(-1)
		  redim HintToTableStart(CountGraphHints,ActiveFileList.Ubound)
		  redim HintToTableEnd(CountGraphHints,ActiveFileList.Ubound)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function isPhraseAnX(inPhrase as String) As Boolean
		  If lowercase(inPhrase .left(1)) = "x" then
		    return true
		  else
		    return false
		  end if
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadAllFiles()
		  dim i as Integer
		  dim curFile as FolderItem
		  
		  For i = 0 to UBound(ActiveFileList)
		    Call ReadHTML(i)
		  next i
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ReadGraphHints()
		  dim SourceStream as TextInputStream
		  dim curLine as string = ""
		  dim hintFolder as FolderItem
		  dim hintFile as FolderItem
		  dim phrases() as string
		  
		  CountGraphHints = -1
		  redim ghReportName(-1)
		  redim ghSubtableName(-1)
		  redim ghCombinedName(-1)
		  redim GraphHints(-1)
		  ' did not seem to always work: 'hintFolder = SpecialFolder.CurrentWorkingDirectory
		  hintFolder = GetFolderItem("")
		  if DebugBuild then
		    hintFolder = hintfolder.Parent 'for debugging purposes only
		  end if
		  hintFile = hintFolder.child("GraphHints.csv")
		  IF hintFile.Exists then
		    SourceStream = hintFile.OpenAsTextFile
		    'skip the first two lines since they are headers
		    if not SourceStream.EOF then curLine = SourceStream.ReadLine
		    if not SourceStream.EOF then curLine = SourceStream.ReadLine
		    While Not SourceStream.EOF
		      curLine = SourceStream.ReadLine
		      phrases = Split(curLine,",")
		      if ubound(phrases) = 18 then
		        CountGraphHints = CountGraphHints + 1
		        redim ghReportName(CountGraphHints)
		        ghReportName(CountGraphHints) = phrases(0)
		        redim ghSubtableName(CountGraphHints)
		        ghSubtableName(CountGraphHints) =  phrases(1)
		        redim ghCombinedName(CountGraphHints)
		        if phrases(1) = "" then
		          ghCombinedName(CountGraphHints) = phrases(0)
		        else
		          ghCombinedName(CountGraphHints) = phrases(0) + " --- " + phrases(1)
		        end if
		        If CountGraphHints>GraphHints.ubound then
		          redim GraphHints(CountGraphHints + 100)
		          redim HintToTableStart(CountGraphHints + 100,ActiveFileList.ubound)
		          redim HintToTableEnd(CountGraphHints + 100, ActiveFileList.Ubound)
		        end if
		        'isBarForEveryValue    Show simple bar graphs for every individual value in table
		        GraphHints(CountGraphHints).isBarForEveryValue = isPhraseAnX(phrases(2))
		        'isStackedBarForEachColumn    Stacked bar of values from each column
		        GraphHints(CountGraphHints).isStackedBarForEachColumn = isPhraseAnX(phrases(3))
		        'isStackedBarForEachRow    Stacked bar of values from each row
		        GraphHints(CountGraphHints).isStackedBarForEachRow = isPhraseAnX(phrases(4))
		        'is100StackedBarForEachEachColumn    100% stacked bar of values from each column
		        GraphHints(CountGraphHints).is100StackedBarForEachColumn = isPhraseAnX(phrases(5))
		        'is100StackedBarForEachEachRow    100% stacked bar of values from each row
		        GraphHints(CountGraphHints).is100StackedBarForEachRow = isPhraseAnX(phrases(6))
		        'isSideBySideForEachColumn    Side-by-side bar of values from each column
		        GraphHints(CountGraphHints).isSideBySideBarForEachColumn = isPhraseAnX(phrases(7))
		        'isSideBySideForEachRow    Side-by-side bar of values from each row
		        GraphHints(CountGraphHints).isSideBySideBarForEachRow = isPhraseAnX(phrases(8))
		        'isSideBySideForTotals    Side-by-side bar of each value in total section across instance of report in file
		        GraphHints(CountGraphHints).isSideBySideForTotals = isPhraseAnX(phrases(9))
		        'isMonthlyLineForEachColumn    Monthly line graph for values from each column
		        GraphHints(CountGraphHints).isMonthlyLineForEachColumn = isPhraseAnX(phrases(10))
		        'numBottomRowsToExclude    Number of bottom rows to exclude (when using multiple values)
		        GraphHints(CountGraphHints).numBottomRowsToExclude = phrases(11).Val
		        'numTopRowsToExclude    Number of top rows to exclude (not including label, when using multiple values)
		        GraphHints(CountGraphHints).numTopRowsToExclude = phrases(12).Val
		        'numRightColumnsToExclude    Number of columns at right to exclude (when using multiple values)
		        GraphHints(CountGraphHints).numRightColumnsToExclude = phrases(13).Val
		        
		        'numLeftColumnsToExclude    Number of columns at left to exclude (not including label, when using multiple values)
		        ' CR8115 
		        ' GraphHints(CountGraphHints).numRightColumnsToExclude = phrases(14).Val
		        GraphHints(CountGraphHints).numLeftColumnsToExclude = phrases(14).Val
		        
		        'isUnusualTable    Unusual table (Special rules)
		        GraphHints(CountGraphHints).isUnusualTable = isPhraseAnX(phrases(15))
		        'isVariableRowCount    Variable Row Count
		        GraphHints(CountGraphHints).isVariableRowCount = isPhraseAnX(phrases(16))
		        'isMultipleInstancePerFile    Multiple Instances Per File
		        GraphHints(CountGraphHints).isMultipleInstancesPerFile = isPhraseAnX(phrases(17))
		      else
		        MsgBox "Error reading GraphHints.csv at line: " + curLine + " number of fields: " + str(UBound(phrases))
		        exit while
		      end if
		    wend
		  else
		    MsgBox "Did not find file " + hintFile.AbsolutePath
		  end if
		  SourceStream.Close
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ReadHTML(inFileNumber as integer)
		  dim SourceStream as TextInputStream
		  dim curLine as string = ""
		  dim curReport as string = ""
		  dim curFor as string = ""
		  dim curSubtable as string = ""
		  dim curCell as String = ""
		  dim ColumnCount as Integer = 0
		  dim discard as Integer = 0
		  dim curColCount as Integer = 0
		  dim HTMLfile as FolderItem
		  dim j as Integer
		  dim found as integer
		  dim readMode as Integer = 0
		  const rmInTableFirstLine = 1
		  const rmInTableOtherLines = 2
		  const rmInTOC = 3
		  const rmSearchForTable = 4
		  
		  curLine = ""
		  readMode = rmSearchForTable
		  'retrieve the file name based on the number
		  HTMLfile = new FolderItem(ActiveFileList(inFileNumber))
		  If HTMLfile <> Nil then
		    SourceStream = HTMLfile.OpenAsTextFile
		    While Not SourceStream.EOF
		      curLine = SourceStream.ReadLine
		      'if a horizontal rule <hr> is found in the file it is a new report, reset main variables and switch to searching for a table
		      If curLine.Left(4) = "<hr>" then
		        readMode = rmSearchForTable
		        curReport = ""
		        curFor = ""
		        curSubtable = ""
		      end if
		      Select case (readMode)
		      case rmInTOC
		        ' do nothing - just skip over the table of contents
		      case rmSearchForTable 'looking for the table
		        ' Table of contents
		        if curLine = "<a name=toc></a>" then
		          readMode = rmInTOC
		        elseif curLine.left(13) = "<p>Report:<b>" then 'Report
		          curReport =curLine.mid(15,curLine.Len - 22)
		        elseif curLine.left(10) = "<p>For:<b>" then 'For
		          curFor = curLine.mid(12,curLine.Len - 19)
		        elseif curLine.left(3) = "<b>" then 'subtable
		          if curLine.right(12) = "</b><br><br>" then
		            curSubtable = curLine.mid(4, curLine.Len - 15)
		          end if
		        elseif curLine.left(14) = "<table border=" then 'Table
		          readMode = rmInTableFirstLine
		        elseif curLine = "Values in table are in hours.<br>" then ' this a timebins report - skip it (as if it is another table of contents)
		          readMode = rmInTOC
		        elseif curLine = "<b>Statistics</b><br><br>" then 'this is a timebins report - skip it
		          readMode = rmInTOC
		        end if
		      case rmInTableFirstLine 'in the table on the first line with column headings
		        if curLine = "</table>" then
		          curSubtable = ""
		          readMode = rmSearchForTable
		        elseif curLine = "  </tr>" then
		          readMode = rmInTableOtherLines
		        elseif curLine = "  <tr><td></td>" then
		          ColumnCount = 0
		          CountTableResults = CountTableResults + 1
		          if ubound(TableResults)<CountTableResults then
		            redim TableResults(CountTableResults + 100)
		          end if
		          TableResults(CountTableResults).ReportNameIndex  = AddName(curReport)
		          TableResults(CountTableResults).ForNameIndex = AddName(curFor)
		          TableResults(CountTableResults).SubTableNameIndex = AddName(curSubtable)
		          TableResults(CountTableResults).NumberOfColumns = 0
		          TableResults(CountTableResults).NumberOfRows = 0
		          TableResults(CountTableResults).FirstValueIndex = 0
		          ' associate the current table with the set of hints previously read in from GraphHints.csv
		          TableResults(CountTableResults).GraphHintIndex = 0
		          found = -1
		          For j = 0 to CountGraphHints
		            If lowercase(ghReportName(j)) = lowercase(curReport) then
		              if lowercase(ghSubtableName(j)) = lowercase(curSubtable) then
		                found = j
		                exit for
		              end if
		            end if
		          next j
		          'check if any hint is assoicated and if not add it to the list of hints with default values
		          if found > -1 then 'was found
		            TableResults(CountTableResults).GraphHintIndex = found
		            If HintToTableStart(found,inFileNumber) = 0 then
		              HintToTableStart(found,inFileNumber) = CountTableResults
		            end if
		            HintToTableEnd(found,inFileNumber) = CountTableResults
		          else 'not found
		            CountGraphHints = CountGraphHints + 1
		            If CountGraphHints>GraphHints.ubound then
		              redim GraphHints(CountGraphHints + 100)
		            end if
		            if CountGraphHints > ubound(HintToTableStart,1) then
		              redim HintToTableStart(CountGraphHints + 100,ActiveFileList.ubound)
		              redim HintToTableEnd(CountGraphHints + 100, ActiveFileList.Ubound)
		            end if
		            'assume that the report is a user defined monthly report
		            ghReportName.append curReport
		            ghSubtableName.append curSubtable
		            if curSubtable= "" then
		              ghCombinedName.append curReport
		            else
		              ghCombinedName.append curReport + " --- " + curSubtable
		            end if
		            GraphHints(CountGraphHints).isBarForEveryValue = False
		            GraphHints(CountGraphHints).isStackedBarForEachColumn = False
		            GraphHints(CountGraphHints).isStackedBarForEachRow = False
		            GraphHints(CountGraphHints).is100StackedBarForEachColumn = False
		            GraphHints(CountGraphHints).is100StackedBarForEachRow = False
		            GraphHints(CountGraphHints).isSideBySideBarForEachColumn = False
		            GraphHints(CountGraphHints).isSideBySideBarForEachRow = False
		            GraphHints(CountGraphHints).isSideBySideForTotals = True
		            GraphHints(CountGraphHints).isMonthlyLineForEachColumn = True
		            GraphHints(CountGraphHints).numBottomRowsToExclude = 4
		            GraphHints(CountGraphHints).numTopRowsToExclude = 0
		            GraphHints(CountGraphHints).numRightColumnsToExclude = 0
		            GraphHints(CountGraphHints).numRightColumnsToExclude = 0
		            GraphHints(CountGraphHints).isUnusualTable = False
		            GraphHints(CountGraphHints).isVariableRowCount = False
		            GraphHints(CountGraphHints).isMultipleInstancesPerFile = True
		            TableResults(CountTableResults).GraphHintIndex = CountGraphHints
		            HintToTableStart(CountGraphHints,inFileNumber) = CountTableResults
		            HintToTableEnd(CountGraphHints,inFileNumber) = CountTableResults
		          end if
		        elseif curLine.Mid(5,10) = "<td align=" then
		          'since this is the first line each cell of the table contains column headings
		          curCell = curLine.Mid(23,curLine.Len - 27)
		          If TableResults(CountTableResults).NumberOfColumns = 0 then
		            TableResults(CountTableResults).FirstColHeadNameIndex = AddName(curCell)
		            TableResults(CountTableResults).NumberOfColumns = 1
		          else
		            discard = AddName(curCell)
		            TableResults(CountTableResults).NumberOfColumns =TableResults(CountTableResults).NumberOfColumns + 1
		          end if
		        end if
		      case rmInTableOtherLines 'in the table on remaining lines with values
		        if curLine = "</table>" then
		          curSubtable = ""
		          readMode = rmSearchForTable
		          continue while
		        elseif curLine.mid(3,4) = "<tr>" then 'new row
		          TableResults(CountTableResults).NumberOfRows =  TableResults(CountTableResults).NumberOfRows + 1
		          curColCount = 0
		        elseif curLine.Mid(5,10) = "<td align=" then
		          curCell = curLine.Mid(23,curLine.Len - 27)
		          curColCount = curColCount + 1
		          if curColCount = 1 then
		            if TableResults(CountTableResults).NumberOfRows = 1 then
		              TableResults(CountTableResults).FirstRowHeadNameIndex =  AddName(curCell) 'first row heading
		            else
		              discard = AddName(curCell) 'next row heading since in column 1
		            end if
		          else 'not a row or column heading - it is table value
		            If TableResults(CountTableResults).FirstValueIndex = 0 then
		              TableResults(CountTableResults).FirstValueIndex = AddValueFromString(curCell) ' first table value
		            else
		              discard = AddValueFromString(curCell) 'another table value
		            end if
		          end if
		        end if
		      end select
		    wend
		    SourceStream.Close
		  end if
		  
		  
		  
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ReadOldFileList()
		  dim SourceStream as TextInputStream
		  dim curLine as string = ""
		  dim listFolder as FolderItem
		  dim listFile as FolderItem
		  dim possibleFile as FolderItem
		  
		  listFolder = SpecialFolder.ApplicationData
		  listFile = listFolder.child("EP-Compare-File-List.txt")
		  IF listFile.Exists then
		    SourceStream = listFile.OpenAsTextFile
		    'skip the first two lines since they are headers
		    While Not SourceStream.EOF
		      curLine = SourceStream.ReadLine
		      try
		        possibleFile = new FolderItem(curLine)
		        if possibleFile.Exists then
		          ActiveFileList.Append curLine
		        end if
		      catch err
		        'did not find file or filefolder
		      end try
		    wend
		    SourceStream.Close
		    call UpdateActiveFileTimeStamps
		  end if
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub UpdateActiveFileTimeStamps()
		  dim curFile as FolderItem
		  dim i as integer
		  for i = 0 to ActiveFileList.Ubound
		    curFile = new FolderItem(ActiveFileList(i))
		    ActiveFileTimeStamp.Append curFile.ModificationDate
		  next i
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub WriteOldFileList()
		  dim outStream as TextOutputStream
		  dim curLine as string = ""
		  dim listFolder as FolderItem
		  dim listFile as FolderItem
		  dim i as Integer
		  
		  listFolder = SpecialFolder.ApplicationData
		  listFile = listFolder.child("EP-Compare-File-List.txt")
		  outStream = listFile.CreateTextFile
		  for i = 0 to ActiveFileList.Ubound
		    outStream.WriteLine ActiveFileList(i)
		  next i
		  outStream.Close
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		#tag Note
			Contains the name of the active files.
		#tag EndNote
		ActiveFileList() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		ActiveFileTimeStamp() As Date
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			Number of GraphHints
		#tag EndNote
		CountGraphHints As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			Count of the Nm array
		#tag EndNote
		CountNm As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			Count of the TableResults array
		#tag EndNote
		CountTableResults As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			Cound of the V array
		#tag EndNote
		CountV As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			Holds the selected graph code
		#tag EndNote
		currentGraphCode As String
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			The combined report and subtable names related to the GraphHints array
		#tag EndNote
		ghCombinedName() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			The report names related to the GraphHints array
		#tag EndNote
		ghReportName() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			The subtable names related to the GraphHints array
		#tag EndNote
		ghSubtableName() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			This is the data from the GraphHints.csv file that show hints about how each table should be shown as graphs
		#tag EndNote
		GraphHints() As GraphHintsType
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			The array contains references to the last item in the TableResults array that corresponds to the GraphHint index. It is used with the HintToTableStart array.
			The form is HintToTableEnd(hintIndex, fileIndex) where the number of the hintIndex corresponds directly with the GraphHints and gh arrays.
			The fileIndex corresponds with the ActiveFileList items.
		#tag EndNote
		HintToTableEnd(-1,-1) As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			The array contains references to the first item in the TableResults array that corresponds to the GraphHint index. It is used with the HintToTableEnd array.
			The form is HintToTableStart(hintIndex, fileIndex) where the number of the hintIndex corresponds directly with the GraphHints and gh arrays.
			The fileIndex corresponds with the ActiveFileList items.
		#tag EndNote
		HintToTableStart(-1,-1) As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			Holds the name of the reports, the "fors", the subtables, the column headers, and the row headers. See TableResults array to understand how the values are put in.
		#tag EndNote
		Nm() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			Holds the current chart in a "picture" variable so that it can be easily put in the Clipboard
		#tag EndNote
		pictOfChart As Picture
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			The array holds the indices of the values and names for each instance of a table found in one of the selected HTML files
		#tag EndNote
		TableResults() As TableResultsType
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			Holds the values of the tables as single-precision floating point numbers. See TableResults array to understand how the values are put in.
		#tag EndNote
		v() As Single
	#tag EndProperty


	#tag Structure, Name = GraphHintsType, Flags = &h0, Attributes = \"StructureAlignment \x3D 1"
		isVariableRowCount as Boolean
		  isBarForEveryValue as Boolean
		  isStackedBarForEachColumn as Boolean
		  isStackedBarForEachRow as Boolean
		  is100StackedBarForEachColumn as Boolean
		  is100StackedBarForEachRow as Boolean
		  isSideBySideBarForEachColumn as Boolean
		  isSideBySideBarForEachRow as Boolean
		  isSideBySideBarForTotalsAcrossInstances as Boolean
		  isMonthlyLineForEachColumn as Boolean
		  numBottomRowsToExclude as Integer
		  numTopRowsToExclude as Integer
		  numRightColumnsToExclude as Integer
		  numLeftColumnsToExclude as Integer
		  isUnusualTable as Boolean
		  isMultipleInstancesPerFile as Boolean
		isSideBySideForTotals as Boolean
	#tag EndStructure

	#tag Structure, Name = TableResultsType, Flags = &h0, Attributes = \"StructureAlignment \x3D 1"
		ReportNameIndex as Integer
		  ForNameIndex as Integer
		  SubTableNameIndex as Integer
		  NumberOfRows as Integer
		  NumberOfColumns as Integer
		  FirstValueIndex as Integer
		  FirstRowHeadNameIndex as Integer
		  FirstColHeadNameIndex as Integer
		  CombinedNameIndex as Integer
		GraphHintIndex as Integer
	#tag EndStructure


	#tag ViewBehavior
		#tag ViewProperty
			Name="CountGraphHints"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CountNm"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CountTableResults"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CountV"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="currentGraphCode"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="pictOfChart"
			Group="Behavior"
			Type="Picture"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
