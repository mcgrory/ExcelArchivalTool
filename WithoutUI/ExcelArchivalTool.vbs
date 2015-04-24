' Created October 20 2014
' University Libraries - University of Minnesota - Twin Cities
' Data Management and Curation Initiative / Data Repository for the University of Minnesota (DRUM)
' Written by John McGrory
' Licensed under GNU GPLv3
'
' This script will prompt the user for a source (either a single Excel file or a directory containing Excel
' files) and for a destination directory (to store the conversion products)
'
' By default, all conversion steps are executed (i.e., raw spreadsheet data exported as CSV, cell formulae
' logged into a TXT file, charts/figures exported as PNG images, and exportation of the workbook as an
' HTML copy). These settings can be changed on line 356 of this script. 
'
'


' Function to remove any Windows-defined illegal path characters
Private Function cleanPath(path)
	cleanPath = Replace(path, "<", "")
	cleanPath = Replace(cleanPath, ">", "")
	cleanPath = Replace(cleanPath, ":", "")
	cleanPath = Replace(cleanPath, "\", "")
	cleanPath = Replace(cleanPath, "/", "")
	cleanPath = Replace(cleanPath, "|", "")
	cleanPath = Replace(cleanPath, "?", "")
	cleanPath = Replace(cleanPath, "*", "")
	cleanPath = Replace(cleanPath, chr(34), "")
End Function


FormulaFolderName = "Formulas"
ChartFolderName = "Charts And Figures"
HTMLFolderName = "Visual Representation (HTML Snapshot)"

' Function to create or get subfolder in DestinationFolder (creates folder if it doesn't exist)
'   Creates subfolder within subfolder to correspond to workbook
Private Function folderCheck(bookName,folderName)
	Dim f
	folderPath = DestinationFolder & "\" & folderName
	If Not(fso.FolderExists(folderPath)) Then
		fso.CreateFolder(folderPath)
	End If

	If Not(fso.FolderExists(folderPath & "\" & bookName)) Then
		Set f = fso.CreateFolder(folderPath & "\" & bookName)
	Else
		Set f = fso.GetFolder(folderPath & "\" & bookName)
	End If

	folderCheck = f.path
End Function


' Function to extract charts/figures from given workbook; also creates text log of all data
'   ranges used to create the charts/figures
Private Function extractCharts(wkbook,fileBaseName,fileExtension)
	copyPath = DestinationFolder & "\" & fileBasename & "_copy" & "." & fileExtension
	wkbook.SaveCopyAs copyPath
	Set currBook = objExcel.Workbooks.Open(copyPath)
	For each wksheet in currBook.Worksheets
		usedColumn = wksheet.UsedRange.Column
		If usedColumn > 1 Then
			column = usedColumn - 1
			Set colDeleteRange = wksheet.Range("A1:" & ConvertToLetter(column) & "1")
			For i = colDeleteRange.Columns.Count To 1 Step -1
				colDeleteRange.Columns(i).EntireColumn.Delete
			Next
		End If 

		usedRow = wksheet.UsedRange.Row
		If usedRow > 1 Then
			row = usedRow - 1
			Set rowDeleteRange = wksheet.Range("A1:A" & row)
			For j = rowDeleteRange.Rows.Count to 1 Step -1
				rowDeleteRange.Rows(j).EntireRow.Delete
			Next
		End If

		If wksheet.ChartObjects.Count > 0 Then
			' Make sure image folder exists for charts/figures; create one if it doesn't
			imageFolder = folderCheck(fileBaseName,ChartFolderName)

			' Create a text file to hold all chart/figure ranges
			rangeFilePath = imageFolder & "\" & fileBaseName & "_Ranges.txt"
			ForAppending = 8
			CreateIfNonExistent = True
			Set rangeFile = fso.OpenTextFile(rangeFilePath, ForAppending, CreateIfNonExistent)

			' Loop through charts on current sheet, exporting each as .png file
			chartcounter = 1
			For each objChart in wksheet.ChartObjects
				Set currChart = objChart.Chart
				imagePath = imageFolder & "\" & cleanPath(wksheet.Name) & "_" & chartcounter & ".png"

				rangeFile.WriteLine("Figure Range(s) for " & cleanPath(wksheet.Name) & "_" & chartcounter & ":")		
	
				For each ser in currChart.SeriesCollection
					rangeFile.WriteLine(ser.FormulaLocal)
				Next

				rangeFile.WriteLine()
	
				objChart.Activate
				currChart.Export imagePath, "PNG"
				chartcounter = chartcounter + 1
			Next		

			rangeFile.Close
			' Counter for report file; substract 1 because update occurs after export
			numberCharts = numberCharts + chartCounter - 1
		End If
	Next

	currBook.Close False
	fso.DeleteFile(copyPath) 

End Function 


' Function to translate cell address to account for deletion of rows/columns when converting to csv
Private Function convertAddress(column,row,rangeUsed,includeAbsolutes)
	firstColumn = rangeUsed.Column
	firstRow = rangeUsed.Row

	If includeAbsolutes Then
		convertAddress = "$" & ConvertToLetter(column - firstColumn + 1) & "$" & (row - firstRow + 1)
	Else
		convertAddress = ConvertToLetter(column - firstColumn + 1) & (row - firstRow + 1)
	End If

End Function


' Function to translate column number into excel style column header letters (e.g., 1=A, 26=Z, 27=AZ, 703=AAA, etc)
Function ConvertToLetter(colNum)
	Dim quotient
	Dim remainder
	If colNum <= 0 Then
		ConvertToLetter = ""
	ElseIf colNum <= 26 Then
		ConvertToLetter = Chr(colNum + 64)
	Else
		quotient = Int(colNum / 26)
		remainder = colNum Mod 26
		If remainder = 0 Then
			quotient = quotient - 1
			remainder = 26
		End If
		ConvertToLetter = ConvertToLetter(quotient) & ConvertToLetter(remainder)
	End If
End Function


' Function to translate formula to account for deletion of unused, leading rows/columns when converting to csv 
'   All absolute references will be converted to relative references
Private Function convertFormula(sheet,curCell)
	' Get first used column and row, calculate these values after csv conversion
	firstColumn = sheet.UsedRange.Column
	firstRow = sheet.UsedRange.Row
	newColumn = curCell.Column - firstColumn + 1
	newRow = curCell.Row - firstRow + 1

	' Rid formula of $ characters and replace with original formula to allow relative copy/paste
	formulaParts = Split(curCell.formula, "$")
	newFormula = Join(formulaParts,"")
	curCell.Value = newFormula

	' Store value in cell that's pasted to; then execute copy/paste
	oldValue = sheet.Cells(newRow,newColumn).Value
	sheet.Range(curCell.Address).Copy sheet.Range(convertAddress(curCell.Column,curCell.Row,sheet.UsedRange,True))

	' Return the converted formula
	convertFormula = sheet.Cells(newRow,newColumn).Formula

	' Replace original value to cell that was pasted to
	sheet.Cells(newRow,newColumn).Value = oldValue
End Function


' Function to grab formulas from worksheet and adjust for deletion of rows/columns on csv conversion
Private Function extractFormulas(wksheet,fileBaseName)
	ForAppending = 8
	CreateIfNonExistent = True
	formulaFolderPath = folderCheck(fileBaseName,FormulaFolderName)
	formulaFilePath = formulaFolderPath & "\" & cleanPath(wksheet.Name) & ".txt"
	Set objFile = fso.OpenTextFile(formulaFilePath, ForAppending, CreateIfNonExistent)

	For each cell in wksheet.UsedRange
		If cell.HasFormula Then
			objFile.WriteLine(convertAddress(cell.Column,cell.Row,wksheet.UsedRange,True) & ": " & convertFormula(wksheet,cell))
		End If													
	Next

	objFile.Close

	' Do empty file/folder checks to avoid having to check folder and open/close file in every loop
	Set file = fso.GetFile(formulaFilePath)
	If file.size <= 0 Then
		fso.DeleteFile file, True
	Else
		numberFormulaFiles = numberFormulaFiles + 1
	End If

	Set folder = fso.GetFolder(formulaFolderPath)
	If folder.Files.Count = 0 Then
		fso.DeleteFolder(folder)
	End If

	Set folder = fso.GetFolder(DestinationFolder & "\" & FormulaFolderName)
	If folder.SubFolders.Count = 0 Then
		fso.DeleteFolder(folder)
	End If
End Function


' Function to create HTML version of workbook 
Private Function extractHTML(wkbook,fileBaseName)
	htmlFolderPath = folderCheck(fileBaseName,HTMLFolderName)
	wkBook.SaveAs htmlFolderPath & "\" & fileBaseName & ".htm", 44
End Function


' Function to extract raw spreadsheet data from worksheet, convert to CSV
Private Function extractRawData(wksheet,fileBaseName)
	newFileBase = fileBaseName & "_" & cleanPath(wksheet.Name)

	' Checks if worksheet is not empty; saves as csv if not empty
	If Not(wksheet.usedrange.address = "$A$1" And wksheet.Range("A1") = "") Then
		wksheet.SaveAs DestinationFolder & "\" & newFileBase & ".csv", 6
		numberCSVFiles = numberCSVFiles + 1
	End If
End Function


' Function to create report text file at end of program run
Private Function createReport()
	reportFilePath = DestinationFolder & "\ Output Report.txt"
	ForAppending = 8
	CreateIfNonExistent = True
	Set reportFile = fso.OpenTextFile(reportFilePath, ForAppending, CreateIfNonExistent)

	reportFile.WriteLine("Excel Archival Tool Conversion Report")
	reportFile.WriteLine("Source: " & Source)
	reportFile.WriteLine("Completed: " & Date & " " & Time)
	reportFile.WriteLine("")
	reportFile.WriteLine("Number of Excel Files processed: " & numberXLSFiles)

	If getData Then
		reportFile.WriteLine("Number of CSV files generated: " & numberCSVFiles)
	End If

	If getChartsFigures Then
		reportFile.WriteLine("Number of charts/figures exported: " & numberCharts)
	End If

	If getFormulas Then
		reportFile.WriteLine("Number of formula files generated: " & numberFormulaFiles)
	End If

	reportFile.Close
End Function


' Function to capture source and destination folders from user
Private Function displaydialogs()

	' Select a single Excel file or a folder that contains Excel files
	myStartLocation = "C:\Users"
	blnSimpleDialong = False

	Const WINDOW_HANDLE = 0  'Must ALWAYS be 0

	Dim numOptions, objFolder, objFolderItem
	Dim objPath, objShell, strPath, strPrompt

	' Set the options for the dialog window
	strPrompt = "Select single Excel file or a directory containing Excel files:"

	' Option = 16384 allows selection of a single file or a folder
	numOptions = 16384  
    
	' Create a Windows Shell object
	Set objShell = CreateObject( "Shell.Application" )  

	strPath = myStartLocation

	' Enter strPath (myStartLocation) as fourth argument if you'd like to be prompted from a different folder (default is desktop)
	Set objFolder = objShell.BrowseForFolder( WINDOW_HANDLE, strPrompt, _
                                              numOptions )

	' Quit if no folder was selected
	If objFolder Is Nothing Then
    		Source = ""
    		Wscript.Echo "Exiting script. No source folder was selected"
    		Wscript.Quit
	End If

	' Retrieve the path of the selected folder and assign to Source
	Set objFolderItem = objFolder.Self
	Source = objFolderItem.Path


	strPrompt = "Select a destination directory (This is where the conversion products will go):"
	numOptions = 0  'Simple dialog (can only select directories)

	' Enter strPath (myStartLocation) as fourth argument if you'd like to be prompted from a different folder (default is desktop)
	Set objFolder = objShell.BrowseForFolder( WINDOW_HANDLE, strPrompt, _
						      numOptions )

	' Quit if no folder was selected
	If objFolder Is Nothing Then
    		Source = ""
    		Wscript.Echo "Exiting Script. No Destination folder was selected"
    		Wscript.Quit
	End If

	' Retrieve the path of the selected folder and assign to DestinationFolder
	Set objFolderItem = objFolder.Self
	DestinationFolder = objFolderItem.Path
End Function


' Function to recursively count number of Excel files in selected directory
Private Function getCount(folder)
	Set folder = fso.GetFolder(folder)
	' Count how many excel files in source folder and give time estimate for completion
	count = 0
	For Each file in folder.files
		If LCase(fso.GetExtensionName(file)) = "xls" Or LCase(fso.GetExtensionName(file)) = "xlsx" Then
			count = count + 1
		End If
	Next

	timeEst = count 

	If count = 1 Then
		verb = "is"
		fileWord = "file"
	Else
		verb = "are"
		fileWord = "files"
	End If 

	response = Msgbox("There " & verb & " " & count & " Excel " & fileWord & " in this folder. Conversion may take up to "_
				 & timeEst & " minute(s) to complete. Do you want to proceed?", vbYesNo)

	If response = vbNo Then
		Wscript.Quit
	End If

End Function


' Initiate and assign File/Folder System objects
Set fso = CreateObject("Scripting.FileSystemObject")


' Initiate an Excel application; set alert display and visibility to false (runs in background)
Set objExcel = CreateObject("Excel.Application")
objExcel.DisplayAlerts = False                         
objExcel.Visible = False


' Flags to mark which actions are desired
getData = True
getFormulas = True
getChartsFigures = True
getHTMLRepresentation = True
					

' Data for output report
numberXLSFiles = 0
numberCSVFiles = 0
numberCharts = 0
numberFormulaFiles = 0


'Function containing main logic of program
Private Function convertExcel(file)
	If (LCase(fso.GetExtensionName(file)) = "xls" Or LCase(fso.GetExtensionName(file)) = "xlsx") Then
	
		' Cycle through each worksheet
		Set oBook = objExcel.Workbooks.Open(file)
		
		' Get file name without .xls or .xlsx extension
		fileBase = cleanPath(fso.GetBaseName(file))
		fileExtension = fso.GetExtensionName(file)

		If getChartsFigures Then
			Call extractCharts(oBook,fileBase,fileExtension)
		End If

		If getHTMLRepresentation Then
			Call extractHTML(oBook,fileBase)
		End If

		For each worksheet in oBook.Worksheets

			If getFormulas Then
				Call extractFormulas(worksheet,fileBase)
			End If

	
			If getData Then
				Call extractRawData(worksheet,fileBase)
			End If

		Next
		
		' Close workbook when finished
		oBook.Close False

		numberXLSFiles = numberXLSFiles + 1

	End If	

End Function


' Function for recursive folder conversion if a folder is selected
Private Function convertFolder(folderPath)

	Set folder = fso.GetFolder(folderPath)
	Set currFiles = folder.Files
	
	' Cycle through each file and check if it's an excel file
	For each file in currFiles
		Call convertExcel(file)
	Next
	

	Set currFolder = fso.GetFolder(folderPath)
	For each subfolder in currFolder.SubFolders
		convertFolder(subfolder.Path)
	Next

End Function


' Function to determine if file or folder is selected and call the appropriate conversion function
Private Function determineInput(Source)
	If fso.FolderExists(Source) Then
		Call getCount(Source)
		Call convertFolder(Source)
	ElseIf fso.FileExists(Source) Then
		Call convertExcel(Source)
	End If
End Function

Source = ""
DestinationFolder = ""
Call displayDialogs()
Wscript.Echo Source
Wscript.Echo DestinationFolder


Call determineInput(Source)


' Close Excel application
objExcel.Quit


Call createReport()


WScript.Echo "Conversion complete."
