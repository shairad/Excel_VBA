Sub Format_Sheets_As_Tables()
'Best used when you want to include all data stored on the spreadsheet

	Dim sht As Worksheet
	Dim LastRow As Long
	Dim LastColumn As Long
	Dim StartCell As Range
	Dim Sheet As Worksheet
	Dim Sheet_Name As String

'Adjusts color theme for the new 2010 colors
	ActiveWorkbook.Theme.ThemeColorScheme.Load ( _
	"C:\Program Files\Microsoft Office\Root\Document Themes 16\Theme Colors\Office 2007 - 2010.xml" _
	)

	For each Sheet in Worksheets 'Loop for each sheet in the workbook

		Sheet.Activate

		Set sht = Sheet 'Sets value
		Set StartCell = Range("A1") 'Start cell used to determine where to begin creating the table range

'Find Last Row and Column
		LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
		LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column
		Sheet_Name = sheet 'Assigns sheet name to a variable as a string

'Select Range
		sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

'Creates the table
		Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
		tbl.Name = Sheet_Name 'Names the table
		tbl.TableStyle = "TableStyleLight12" 'Sets table color theme
		Next Sheet

	End Sub
