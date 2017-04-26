Sub EV_Code_Setup()
    
    Dim tbl As ListObject    
    Dim sht As Worksheet    
    Dim LastRow As Long    
    Dim LastColumn As Long    
    Dim StartCell As Range    
    Dim PvtTbl As PivotTable
    
    Application.ScreenUpdating = False        
    
    Set sht = Worksheets("Event Codes Results")    
    With sht   ' replaced 'Select' which slows down your code with `With` Statement        
        If .ListObjects.Count > 0 Then            
            With .ListObjects(1)                
                Set rList = .Range                
                .Unlist                           ' convert the table back to a range            
            End With
            With rList                
                .Interior.ColorIndex = xlColorIndexNone                
                .Font.ColorIndex = xlColorIndexAutomatic                
                .Borders.LineStyle = xlLineStyleNone            
            End With  
        End If                
        Set StartCell = .Range("A1")
        'Refresh UsedRange        
    '   .UsedRange ' *** <-- Not sure what you are trying to achieve with this line ?     

        'Find Last Row and Column        
        LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row        
        LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column            

        ' Edited : Turn Range Into Table (without Selection)        
        Set tbl = .ListObjects.Add(xlSrcRange, .Range(StartCell, .Cells(LastRow, LastColumn)), , xlYes)        
        tbl.Name = "EV_Results_Table"        
        tbl.TableStyle = "TableStyleLight9"            
    
        'changes font color of header row to white        
        With .Rows("1:1").Font            
            .ThemeColor = xlThemeColorDark1           
            .TintAndShade = 0        
        End With            

        .Range("A2") = "=IFERROR(INDEX('Validated Codes'!I:I,MATCH(D3,'Validated Codes'!D:D,0)),0)"        
        .Range("A2").AutoFill Destination:=Range("EV_Results_Table[Mapped?]")    
    End With

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _   
        "EV_Results_Table", Version:=6).CreatePivotTable TableDestination:= _       
        "'Pivot Table'!R12C2", TableName:="EV_Pivot", DefaultVersion:=6       

    ' set the Pivot Table "EV_Pilot" to an Object   
    Set PvtTbl = Sheets("Pivot Table").PivotTables("EV_Pilot")   
    With PvtTbl       
        With .PivotFields("Mapped?")        
            .Orientation = xlRowField           
            .Position = 1        
        End With        
        With .PivotFields("CODE_STATUS")            
            .Orientation = xlColumnField           
            .Position = 1        
        End With       
        ' add field to Pivot Table as Count        
        .AddDataField .PivotFields("EVENT_CD"), "Count of EVENT_CD", xlCount    
    End With
    
    'Selects the Validated count and changes color to red    
    With Sheets("Pivot Table").Range("C15").Font        
        .Color = -16776961        
        .TintAndShade = 0    
    End With

    With tbl ' <-- you already defined it, so you can use this Object        
        'Filters the results table column A for just "Validated"       
        .Range.AutoFilter Field:=1, Criteria1:="Validated"                'Filters the Code_Status column for just "Active"       
        .Range.AutoFilter Field:=3, Criteria1:="Active"    
    End With
    Application.ScreenUpdating = True

End Sub
