Sub Change_Colors()
'
' Changes the workbook color theme to the 2007 - 2010 colors.
'

	ActiveWorkbook.Theme.ThemeColorScheme.Load ( _
	"C:\Program Files\Microsoft Office\Root\Document Themes 16\Theme Colors\Office 2007 - 2010.xml" _
	)
End Sub
