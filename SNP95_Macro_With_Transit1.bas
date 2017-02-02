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
        
            
    ' Number rows for sorting, leaving space for row 7 to be populated later
        Dim lastrow As Long
         With ActiveSheet
            lastrow = .Cells(.Rows.Count, "E").End(xlUp).Row
            End With
            
        For Each cell In Range("E2:E7")
            cell.Value = cell.Row - 1
        Next
        For Each cell In Range("E8:E10")
            cell.Value = cell.Row
        Next
        For Each cell In Range("E11:E" & lastrow)
            cell.FormulaR1C1 = "=R[-9]C"
            cell.Value = cell.Value
        Next
          
    ' Create Tactical Planning rows
        ActiveSheet.UsedRange.AutoFilter Field:=5, Criteria1:="4"
    
       
        Range("A5:F5", Range("A5:F5").End(xlDown)).Copy
    
        Range("A5:F5").End(xlDown).Offset(6, 0).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        ActiveCell.Offset(0, 4) = "7"
        ActiveCell.Offset(0, 5) = "Tactical Planning"
        ActiveCell.Offset(0, 4).Select
        Range(Selection, Selection.End(xlToRight)).Copy
        Range(Selection, Selection.End(xlDown)).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
    
    ' Create stock in transit rows
        Cells.AutoFilter Field:=5, Criteria1:="7"
    
        Range("A1").Offset(4, 0).Range("A1:F1").Select
        Range(Selection, Selection.End(xlDown)).Copy
    
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Range("A1").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        ActiveCell.Offset(0, 4).Value = "11"
        ActiveCell.Offset(0, 5).Value = "In Transit"
        ActiveCell.Offset(0, 4).Select
        Range(Selection, Selection.End(xlToRight)).Copy
    
        Range(Selection, Selection.End(xlDown)).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
    
        Cells.AutoFilter
    
        
    ' Sorts product via cntry via order of Key figures
        Cells.Select
        Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Key2:=Range("d2") _
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
        
        Formula = a & b & C & d & e & f & g & h & i & j & k & l & m & n & o & p
        
        Debug.Print Formula
    
    'define last column in worksheet
        Dim lastColumn As Long
        lastColumn = Range("A1").End(xlToRight).column
    
    'define last row in worksheet
        Dim lastRowNumber As Long
        lastRowNumber = Range("A1").End(xlDown).Row
    

'CALCULATE STOCK IN TRANSIT

    'Input stock in transit formula into stock in transit cells
        Range("H12", Cells(12, lastColumn)).FormulaR1C1 = _
            "=IFERROR(R[-3]C-(R[-3]C[-1]-R[-10]C-R[-9]C-R[-8]C-R[-7]C+R[-6]C+R[-5]C),0)"
   
    'Copy stock in transit formula to all stock in transit rows
        ActiveSheet.UsedRange.AutoFilter Field:=6, Criteria1:= _
            "In Transit"
    
        Range("H12", Range("H12").End(xlToRight)).Copy

        Range("H12", Cells(lastRowNumber, lastColumn)).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        Cells.AutoFilter Field:=6


    'Copy and paste all as values only
        Range("A1").Select
        Cells.Copy
        Cells.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False

'INSERT FORMULAE FOR CALCULATED FIELDS

    'Input stock on hand forumla
        Range("H9", Cells(9, lastColumn)).FormulaR1C1 = "=IF(R[-1]C="""",RC[-1]+R[3]C-SUM(R[-7]C:R[-4]C)+SUM(R[-3]C:R[-2]C),RC[-1]+R[3]C-SUM(R[-7]C:R[-4]C)+SUM(R[-2]C:R[-1]C))"
    

    'Copy stock on hand formula to all stock on hand rows
        Cells.AutoFilter Field:=6, Criteria1:= _
            "Stock on hand(proj.)"
            
        Range("H9", Range("H9").End(xlToRight)).Copy
                
        Range("H9", Cells(lastRowNumber, lastColumn)).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        Cells.AutoFilter Field:=6
    
    
    'Input weeks cover forumla
        Range("H11", Cells(11, lastColumn)).FormulaR1C1 = Formula

        
    'Copy weeks cover formula to all weeks cover rows
            'note: this is calculation intensive - adding line to temporarily set calcs to manual
                Application.Calculation = xlCalculationManual
    
        Cells.AutoFilter Field:=6, Criteria1:= _
        "weeks Cover"
        
        Range("H11").Copy
        Range("H11", Cells(lastRowNumber, lastColumn)).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
            Cells.AutoFilter Field:=6
        
           
            'calculations are set back to automatic
                Application.Calculation = xlCalculationAutomatic


' END OF CALCULATIONS. ALL FORMATTING BELOW THIS LINE


' Format first location product
    Range("A2", Cells(12, lastColumn)).Select
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlHairline
    End With
    
    'format first location product numbers
    
        'format numbers G2 thorough G12
            Range("G2", Cells(12, lastColumn)).NumberFormat = "#,##0"
        
        'format weeks cover
            Range("G11", Cells(11, lastColumn)).NumberFormat = "#,##0.0"
        
      
    'select A1
    Range("A1").Select
    
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
    

'Production formatting

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
    
    
    Application.GoTo Reference:="R1C1"
    
'Negative formatting for stock on hand and weeks cover
Application.GoTo Reference:="R9C7:R9C34,R11C7:R11C34"
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = RGB(156, 0, 6)
    End With
    
    With Selection.FormatConditions(1).Interior
        .Color = RGB(255, 199, 206)
    End With
    
    Selection.FormatConditions(1).StopIfTrue = False
    Application.GoTo Reference:="R1C1"
    
'Copy all formatting
    Application.GoTo Reference:="R2C1"
    ActiveCell.Range("A1:A11").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.End(xlUp).Select
    'selects A2 again, then replaces zeroes
    ActiveCell.Offset(1, 0).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:="0", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Application.GoTo Reference:="R1C1"
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
       
    'insert week numbers and format cells
        With Range("H1", Cells(1, lastColumn))
            'insert formula for week numbers
            .FormulaR1C1 = _
                "=IF(""Wk ""&WEEKNUM(SUBSTITUTE(R[1]C,""."",""/""))=""Wk 53"", ""Wk 1"",""Wk ""&WEEKNUM(SUBSTITUTE(R[1]C,""."",""/"")))"
            'format font
                With .Font
                    .Bold = True
                    .Name = "Calibri"
                    .FontStyle = "Bold"
                    .Size = 11
                    .ThemeColor = xlThemeColorDark1
                    .ThemeFont = xlThemeFontMinor
                End With
            'format borders
                With .Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
            'format everything else
                .Interior.Color = 2162853
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
                
        End With

    'hide counters in column E
        Range("E1").EntireColumn.Hidden = True
    
    'add filters
        Application.GoTo Reference:="R1C1"
        Selection.AutoFilter


    'freeze panes
    Application.GoTo Reference:="R3C7"
    ActiveWindow.FreezePanes = True
    
    
    'doesn't this clear all of the formatting that we did earlier?
    Sheet1.Range("A1:AH6000").FormatConditions.Delete

    'increase lastrownumber by 1 to account for row added earlier
        lastRowNumber = lastRowNumber + 1
    
    
    'APPLY CONDITIONAL FORMATTING TO the grid
        Dim grid As Range
        Set grid = Range("H3", Cells(lastRowNumber, lastColumn))

        'Add formatting for Stock on Hand less than zero
            With grid
                .FormatConditions(1).StopIfTrue = False
                .FormatConditions.Add Type:=xlExpression, Formula1:= _
                    "=IF(AND($F1=""Stock on hand(proj.)"",H1<0),TRUE)"
                .FormatConditions(grid.FormatConditions.Count).SetFirstPriority
                .FormatConditions(1).Font.Color = RGB(156, 0, 6)
                .FormatConditions(1).Interior.Color = RGB(255, 199, 206)
            
        'Add formatting for Stock on hand below safety stock

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
                
        'Add formatting for Something else!
            
                .FormatConditions.Add Type:=xlExpression, Formula1:= _
                    "=IF(AND($F1=""Stock on hand(proj.)"",H1=""""),TRUE)"
                .FormatConditions(grid.FormatConditions.Count).SetFirstPriority
                .FormatConditions(1).Interior.ThemeColor = xlThemeColorDark1
                .FormatConditions(1).StopIfTrue = False
            End With
            
    'insert selection box for overstock and understock criteria
        Range("AL1") = "Overstock - % of Safety Stock(SNP)"
        Range("AL2") = "Understock - % of Safety Stock(SNP)"
        Range("AM1:AM2").Style = "Percent"
        Columns("AL:AL").EntireColumn.AutoFit
    
        Range("AM1") = 1.5
        Range("AM2") = 0.75
        Range("AL1:AM2").Interior.Color = RGB(149, 179, 215)
    
   'Select cell A1 and set zoom to 85
        Range("A1").Select
        ActiveWindow.Zoom = 85
    
End Sub
