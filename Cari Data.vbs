Dim fso, files, file, excelApp
Dim sourceWorkbook, sourceSheet, destinationWorkbook, destinationSheet
Dim searchRange, sourceRange, destinationRange, lastRow
Dim searchWord, foundCell, flag

' Flag for found or not
flag = False

' Prompt the user to enter the search word
searchWord = InputBox("Data yang ingin dicari:", "Cari Data")

' Exit if the user doesn't enter anything
If searchWord = "" Then
    MsgBox "Data kosong. Mengakhiri pencarian."
    WScript.Quit
End If

' Initialize FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Get the path to the folder where the script is located
scriptPath = WScript.ScriptFullName
folderPath = fso.GetParentFolderName(scriptPath) & "\"

' Get the folder object
Set folder = fso.GetFolder(folderPath)
Set files = folder.Files

' Create Excel application object
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False ' Hide Excel window

' Create destination workbook
Set destinationWorkbook = excelApp.Workbooks.Add
Set destinationSheet = destinationWorkbook.Worksheets.Add
destinationSheet.Name = "Search Results"

' Loop through each file in the folder
For Each file In files
	' Check if the file is an Excel file and does not start with ~$
    If (LCase(fso.GetExtensionName(file)) = "xls" Or LCase(fso.GetExtensionName(file)) = "xlsx") And Left(File.Name, 2) <> "~$" Then
		' Open the Excel workbook
        Set sourceWorkbook = excelApp.Workbooks.Open(file.Path)
		
		' Loop through each worksheet in the workbook
        For Each sourceSheet In sourceWorkbook.Sheets
			' Restrict search only from A to D
			Set searchRange = sourceSheet.Range("A:D")
			
			' Search for the word in the current worksheet
            Set foundCell = searchRange.Find(searchWord)
			
			' If find searchWord
			If Not foundCell Is Nothing Then
				' Select A-Q row in foundCell
				Set sourceRange = sourceSheet.Range("A" & foundCell.Row & ":Q" & foundCell.Row)
				
				' Locate last blank row in destinationSheet
				lastRow = destinationSheet.Cells(destinationSheet.Rows.Count, 1).End(-4162).Row + 1 ' -4162 is xlUp
				
				' Copying value from sourceRange to destinationRange
				Set destinationRange = destinationSheet.Cells(lastRow, 1).Resize(1, sourceRange.Columns.Count)
				destinationRange.Value = sourceRange.Value
				
				flag = True
				Set sourceRange = Nothing
				Set destinationRange = Nothing
			End If
			
			Set foundCell = Nothing
		Next
		
		sourceWorkbook.Close False
		Set sourceWorkbook = Nothing
	End If
Next

' Display a message when done
If flag Then
	excelApp.Visible = True ' Show Excel window
	MsgBox "Pencarian Selesai!"
Else
	MsgBox "Tidak ada data yang ditemukan!"
	destinationWorkbook.Close False
	Set destinationWorkbook = Nothing
	excelApp.Quit
	Set excelApp = Nothing
End If

Set fso = Nothing