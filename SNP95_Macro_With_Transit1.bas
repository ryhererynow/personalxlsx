Sub SNP95_colour_2017_V4()

' SNP95 macro - original by Sam Rabi

' 2017 refresh by Ryan Flynn

'INITIAL FORMATTING AND SORTING

    'find location column, then sort by location
        Dim si As Integer
        Dim sii As Integer
        Dim siii As Boolean
        si = 1
        sii = 1
        siii = False
        Do Until si > 10 Or siii = True
            Do Until sii > 10 Or siii = True
                If Left(Cells(si, sii).Value, 8) = "Location" Then
                    
                    Rows(si).AutoFilter
                    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Cells(si, sii), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                With ActiveWorkbook.ActiveSheet.AutoFilter.Sort
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                            
                    Selection.AutoFilter
                    
                    siii = True
                End If
            sii = sii + 1
            Loop
        si = si + 1
        Loop
    
    ' Rename worksheet based on input box
        Set wks = ActiveSheet
        Do While sName <> wks.Name
            sName = Application.InputBox _
              (Prompt:="Enter new worksheet name")
            On Error Resume Next
            wks.Name = sName
            On Error GoTo 0
        Loop
        Set wks = Nothing
        
    ' Delete columns not needed for SNP95
        Union(Range("B:B"), Range("F:F")).Delete
    
    'define last row in worksheet
        Dim lastRowNumber As Long
        lastRowNumber = Range("A1").End(xlDown).Row
    
    'define last column in worksheet
        Dim lastColumn As Long
        lastColumn = Range("A1").End(xlToRight).column
        
    ' Number rows for sorting, leaving space for row 7 to be populated later
    
        For Each cell In Range("E2:E7")
            cell.Value = cell.Row - 1
        Next
        For Each cell In Range("E8:E10")
            cell.Value = cell.Row
        Next
        For Each cell In Range("E11:E" & lastRowNumber)
            cell.FormulaR1C1 = "=R[-9]C"
            cell.Value = cell.Value
        Next
           
    ' Create Tactical Planning rows
        For Each cell In Range("E1:E" & lastRowNumber)
            If cell.Value = "4" Then
            Rows(cell.Row).EntireRow.Copy
            Rows(cell.Row + 3).EntireRow.Insert Shift:=xlDown
            Cells(cell.Row + 3, cell.column).Value = 7
            Cells(cell.Row + 3, cell.column + 1).Value = "Tactical Planning"
            'increase last row number due to added row
            lastRowNumber = lastRowNumber + 1
            
            End If
                        
        Next
        
    'create stock in transit rows
         For Each cell In Range("E1:E" & lastRowNumber)
            If cell.Value = "7" Then
            Rows(cell.Row).EntireRow.Copy
            Rows(cell.Row + 4).EntireRow.Insert Shift:=xlDown
            Cells(cell.Row + 4, cell.column).Value = 11
            Cells(cell.Row + 4, cell.column + 1).Value = "In Transit"
            'increase last row number due to added row
            lastRowNumber = lastRowNumber + 1
            
            End If
                        
        Next
    
    ' Sorts product via country via order of Key figures
        Cells.Sort Key1:=Range("A2"), Order1:=xlAscending, Key2:=Range("d2") _
            , Order2:=xlAscending, Key3:=Range("e2") _
            , Order3:=xlAscending, Header:=xlYes, OrderCustom:=1, MatchCase:= _
            False, Orientation:=xlTopToBottom
        
'DEFINE VARIABLES FOR CALCULATIONS
    'DEFINE FORMULA for weeks cover
        a = "=IFERROR(IF(R[-2]C=0,"""",IF(R[-2]C-SUM(R[-9]C[1]:R[-8]C[1],R[-6]C[1])<0,R[-2]C/SUM(R[-9]C[1]:R[-8]C[1],R[-6]C[1]),"
        b = "IF(R[-2]C-SUM(R[-9]C[1]:R[-8]C[2],R[-6]C[1]:R[-6]C[2])<0,1+(R[-2]C-SUM(R[-9]C[1]:R[-8]C[1],R[-6]C[1]))/SUM(R[-9]C[2]:R[-8]C[2],R[-6]C[2]),"
        C = "IF(R[-2]C-SUM(R[-9]C[1]:R[-8]C[3],R[-6]C[1]:R[-6]C[3])<0,2+(R[-2]C-SUM(R[-9]C[1]:R[-8]C[2],R[-6]C[1]:R[-6]C[2]))/SUM(R[-9]C[3]:R[-8]C[3],R[-6]C[3]),"
        d = "IF(R[-2]C-SUM(R[-9]C[1]:R[-8]C[4],R[-6]C[1]:R[-6]C[4])<0,3+(R[-2]C-SUM(R[-9]C[1]:R[-8]C[3],R[-6]C[1]:R[-6]C[3]))/SUM(R[-9]C[4]:R[-8]C[4],R[-6]C[4]),"
        e = "IF(R[-2]C-SUM(R[-9]C[1]:R[-8]C[5],R[-6]C[1]:R[-6]C[5])<0,4+(R[-2]C-SUM(R[-9]C[1]:R[-8]C[4],R[-6]C[1]:R[-6]C[4]))/SUM(R[-9]C[5]:R[-8]C[5],R[-6]C[5]),"
        f = "IF(R[-2]C-SUM(R[-9]C[1]:R[-8]C[6],R[-6]C[1]:R[-6]C[6])<0,5+(R[-2]C-SUM(R[-9]C[1]:R[-8]C[5],R[-6]C[1]:R[-6]C[5]))/SUM(R[-9]C[6]:R[-8]C[6],R[-6]C[6]),"
        g = "IF(R[-2]C-SUM(R[-9]C[1]:R[-8]C[7],R[-6]C[1]:R[-6]C[7])<0,6+(R[-2]C-SUM(R[-9]C[1]:R[-8]C[6],R[-6]C[1]:R[-6]C[6]))/SUM(R[-9]C[7]:R[-8]C[7],R[-6]C[7]),"
        h = "IF(R[-2]C-SUM(R[-9]C[1]:R[-8]C[8],R[-6]C[1]:R[-6]C[8])<0,7+(R[-2]C-SUM(R[-9]C[1]:R[-8]C[7],R[-6]C[1]:R[-6]C[7]))/SUM(R[-9]C[8]:R[-8]C[8],R[-6]C[8]),"
        i = "IF(R[-2]C-SUM(R[-9]C[1]:R[-8]C[9],R[-6]C[1]:R[-6]C[9])<0,8+(R[-2]C-SUM(R[-9]C[1]:R[-8]C[8],R[-6]C[1]:R[-6]C[8]))/SUM(R[-9]C[9]:R[-8]C[9],R[-6]C[9]),"
        j = "IF(R[-2]C-SUM(R[-9]C[1]:R[-8]C[10],R[-6]C[1]:R[-6]C[10])<0,9+(R[-2]C-SUM(R[-9]C[1]:R[-8]C[9],R[-6]C[1]:R[-6]C[9]))/SUM(R[-9]C[10]:R[-8]C[10],R[-6]C[10]),"
        k = "IF(R[-2]C-SUM(R[-9]C[1]:R[-8]C[11],R[-6]C[1]:R[-6]C[11])<0,10+(R[-2]C-SUM(R[-9]C[1]:R[-8]C[10],R[-6]C[1]:R[-6]C[10]))/SUM(R[-9]C[11]:R[-8]C[11],R[-6]C[11]),"
        l = "IF(R[-2]C-SUM(R[-9]C[1]:R[-8]C[12],R[-6]C[1]:R[-6]C[12])<0,11+(R[-2]C-SUM(R[-9]C[1]:R[-8]C[11],R[-6]C[1]:R[-6]C[11]))/SUM(R[-9]C[12]:R[-8]C[12],R[-6]C[12]),"
        m = "IF(R[-2]C-SUM(R[-9]C[1]:R[-8]C[13],R[-6]C[1]:R[-6]C[13])<0,12+(R[-2]C-SUM(R[-9]C[1]:R[-8]C[12],R[-6]C[1]:R[-6]C[12]))/SUM(R[-9]C[13]:R[-8]C[13],R[-6]C[13]),"
        n = "IF(R[-2]C-SUM(R[-9]C[1]:R[-8]C[14],R[-6]C[1]:R[-6]C[14])<0,13+(R[-2]C-SUM(R[-9]C[1]:R[-8]C[13],R[-6]C[1]:R[-6]C[13]))/SUM(R[-9]C[14]:R[-8]C[14],R[-6]C[14]),"
        o = "IF(R[-2]C-SUM(R[-9]C[1]:R[-8]C[15],R[-6]C[1]:R[-6]C[15])<0,14+(R[-2]C-SUM(R[-9]C[1]:R[-8]C[14],R[-6]C[1]:R[-6]C[14]))/SUM(R[-9]C[15]:R[-8]C[15],R[-6]C[15]),"
        p = "IF(R[-2]C-SUM(R[-9]C[1]:R[-8]C[16],R[-6]C[1]:R[-6]C[16])<0,15+(R[-2]C-SUM(R[-9]C[1]:R[-8]C[15],R[-6]C[1]:R[-6]C[15]))/SUM(R[-9]C[16]:R[-8]C[16],R[-6]C[16]),R[-2]C/AVERAGE(R[-9]C[1]:R[-8]C[17],R[-6]C[1]:R[-6]C[17])))))))))))))))))),"""")"
        
        formula = a & b & C & d & e & f & g & h & i & j & k & l & m & n & o & p
        
        Debug.Print formula
    
'CALCULATE STOCK IN TRANSIT
    For Each cell In Range("F1:F" & lastRowNumber)
        'if cell value = "In Transit", then select all cells 2 columns to the right of the cell. omits final 5 columns due to a
        'mismatch between the planning book calculations and the SNP95 data view, which incorrectly calculates stock in transit
        If cell.Value = "In Transit" Then Range(Cells(cell.Row, cell.column + 2), Cells(cell.Row, lastColumn - 5)).FormulaR1C1 = _
            "=IFERROR(R[-3]C-(R[-3]C[-1]-R[-10]C-R[-9]C-R[-8]C-R[-7]C+R[-6]C+R[-5]C),0)"
    Next
    
    'Copy and paste all as values only
        Cells.Copy
        Cells.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False

    'replace zeroes
       Cells.Replace What:="0", Replacement:="", LookAt:=xlWhole, _
           SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
           ReplaceFormat:=False

'INSERT FORMULAE FOR CALCULATED FIELDS

    'Input stock on hand forumla
        For Each cell In Range("F1:F" & lastRowNumber)
            If cell.Value = "Stock on hand(proj.)" Then Range(Cells(cell.Row, cell.column + 2), Cells(cell.Row, lastColumn)).FormulaR1C1 = _
                   "=IF(R[-1]C="""",RC[-1]+R[3]C-SUM(R[-7]C:R[-4]C)+SUM(R[-3]C:R[-2]C),RC[-1]+R[3]C-SUM(R[-7]C:R[-4]C)+SUM(R[-2]C:R[-1]C))"
    'Input weeks cover forumla
            If cell.Value = "weeks Cover" Then Range(Cells(cell.Row, cell.column + 2), Cells(cell.Row, lastColumn)).FormulaR1C1 = formula
        Next
        
        'autofilter - seems to have an effect on cell width in final sheet.
        Cells.AutoFilter Field:=6, Criteria1:= _
        "weeks Cover"
        Cells.AutoFilter Field:=6

' END OF CALCULATIONS. ALL FORMATTING BELOW THIS LINE

    ' Format first location product
        'format cells for first location product
            With Range("A2", Cells(12, lastColumn))
                .Borders(xlEdgeLeft).Weight = xlMedium
                .Borders(xlEdgeTop).Weight = xlMedium
                .Borders(xlEdgeBottom).Weight = xlMedium
                .Borders(xlEdgeRight).Weight = xlMedium
                .Borders(xlInsideVertical).Weight = xlThin
                .Borders(xlInsideHorizontal).Weight = xlHairline
            End With
        
        'adjust number formatting for first location product key figures
            'format as numbers
                Range("G2", Cells(12, lastColumn)).NumberFormat = "#,##0"
            
            'format weeks cover as numbers with one decimal
                Range("G11", Cells(11, lastColumn)).NumberFormat = "#,##0.0"
        
        'Tactical Planning Formatting
            Dim TP As Range
            Set TP = Range("G8", Cells(8, lastColumn))
            With TP
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
                    Formula1:="="""""
                .FormatConditions(TP.FormatConditions.Count).SetFirstPriority
                .FormatConditions(1).Font.Bold = True
                .FormatConditions(1).Interior.Color = 65535
            End With
            
        'Production Formatting
            Dim production As Range
            Set production = Range("G6", Cells(6, lastColumn))
            With production
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
                    Formula1:="=0"
                .FormatConditions(production.FormatConditions.Count).SetFirstPriority
                .FormatConditions(1).Font.Bold = True
                .FormatConditions(1).Interior.Color = 10092543
                .FormatConditions(1).StopIfTrue = False
            End With
              
        'Negative formatting for stock on hand and weeks cover
            For Each cell In Range("F1:F12")
                If cell.Value = "Stock on hand(proj.)" Or cell.Value = "weeks Cover" Then
                    With Range(Cells(cell.Row, cell.column + 2), Cells(cell.Row, lastColumn))
                        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
                        .FormatConditions(Range(Cells(cell.Row, cell.column + 2), Cells(cell.Row, lastColumn)).FormatConditions.Count).SetFirstPriority
                        .FormatConditions(1).Font.Color = RGB(156, 0, 6)
                        .FormatConditions(1).Interior.Color = RGB(255, 199, 206)
                        .FormatConditions(1).StopIfTrue = False
                    End With
                End If
            Next
    
'Copy all formatting from first location product to all location products
    
    Range("A2", Cells(12, lastColumn)).Copy
    Range("A2", Cells(lastRowNumber, lastColumn)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
        
'Insert headings and final format
    Cells.AutoFilter
    
    'format dates and headers
        With Range("A1", Cells(1, lastColumn))
            With .Interior
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.14996795556505
            End With
            .Font.Bold = True
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeBottom).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
            .Borders(xlInsideVertical).Weight = xlThin
        End With

    'rename columns to reduce column width
        Range("C1").Value = "Cntry"
        Range("D1").Value = "Loc."
        Range("E1").Value = "Ord"

    'resize columns using autofit
        Columns("A:AH").EntireColumn.EntireColumn.AutoFit

    'insert a row above row 1
        Rows("1:1").EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        'increase lastrownumber by 1 to account for row added earlier
        lastRowNumber = lastRowNumber + 1
       
    'insert week numbers and format cells
        With Range("H1", Cells(1, lastColumn))
            'insert formula for week numbers
            .FormulaR1C1 = _
                "=IF(""Wk ""&WEEKNUM(SUBSTITUTE(R[1]C,""."",""/""))=""Wk 53"", ""Wk 1"",""Wk ""&WEEKNUM(SUBSTITUTE(R[1]C,""."",""/"")))"
                With .Font
                    .Bold = True
                    .Name = "Calibri"
                    .Size = 11
                    .Color = RGB(255, 255, 255)
                End With
            'format everything else
                .Interior.Color = RGB(162, 0, 31)
                .Borders(xlInsideVertical).Weight = xlThin
                .HorizontalAlignment = xlCenter
        End With

    'hide counters in column E
        Range("E1").EntireColumn.Hidden = True
    
    'add filters
        Range("A1").AutoFilter

    'freeze panes
    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
            .SplitColumn = 6
            .SplitRow = 2
            .FreezePanes = True
    End With

    'APPLY CONDITIONAL FORMATTING TO the grid for stock on hand
        Dim grid As Range
        Set grid = Range("H3", Cells(lastRowNumber, lastColumn))
            
        'Add formatting for Stock on hand below safety stock
            With grid
                .FormatConditions(1).StopIfTrue = False
                .FormatConditions.Add Type:=xlExpression, Formula1:= _
                    "=IF(AND($F1=""Stock on hand(proj.)"",H1<($AM$2*H2),H1>0),TRUE)"
                .FormatConditions(grid.FormatConditions.Count).SetFirstPriority
                .FormatConditions(1).Interior.Color = RGB(255, 192, 0)
            
        'Add formatting for Stock on hand over safety stock
                .FormatConditions(1).StopIfTrue = False
                .FormatConditions.Add Type:=xlExpression, Formula1:= _
                    "=IF(AND($F1=""Stock on hand(proj.)"",H1>($AM$1*H2)),TRUE)"
                .FormatConditions(grid.FormatConditions.Count).SetFirstPriority
                .FormatConditions(1).Interior.Color = RGB(177, 160, 199)
           
        'Add formatting for Stock on hand within safety stock
                .FormatConditions(1).StopIfTrue = False
                .FormatConditions.Add Type:=xlExpression, Formula1:= _
                    "=IF(AND($F1=""Stock on hand(proj.)"",H1>=($AM$2*H2),H1<=($AM$1*H2)),TRUE)"
                .FormatConditions(grid.FormatConditions.Count).SetFirstPriority
                .FormatConditions(1).Interior.Color = RGB(196, 215, 155)
            End With
            
    'insert selection box for overstock and understock criteria
        Range("AL1") = "Overstock - % of Safety Stock(SNP)"
        Range("AL2") = "Understock - % of Safety Stock(SNP)"
        Columns("AL:AL").EntireColumn.AutoFit
        Range("AM1:AM2").Style = "Percent"
        Range("AM1") = 1.5
        Range("AM2") = 0.75
        Range("AL1:AM2").Interior.Color = RGB(149, 179, 215)
    
   'Set zoom to 85
        ActiveWindow.Zoom = 85
        Range("A3").Select
        Range("A1").Select
    
End Sub
