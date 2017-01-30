Attribute VB_Name = "SNP95_Macro_With_Transit1"
Sub SNP95_colour_2017_V4()

' Sam Rabi SNP95 macro

' 2017 refresh by Ryan Flynn

'Sort by location
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


' Rename worksheet
    Set wks = ActiveSheet
    Do While sName <> wks.Name
        sName = Application.InputBox _
          (Prompt:="Enter new worksheet name")
        On Error Resume Next
        wks.Name = sName
        On Error GoTo 0
    Loop
    Set wks = Nothing
    Application.GoTo Reference:="R1C1"
    
' Delete columns not needed for SNP95
    Union(Range("B:B"), Range("F:F")).Delete
    
        
' Number rows for sorting
    Range("e2").FormulaR1C1 = "1"
    Range("e3").FormulaR1C1 = "2"
    Range("e4").FormulaR1C1 = "3"
    Range("e5").FormulaR1C1 = "4"
    Range("e6").FormulaR1C1 = "5"
    Range("e7").FormulaR1C1 = "6"
    Range("e8").FormulaR1C1 = "8"
    Range("e9").FormulaR1C1 = "9"
    Range("e10").FormulaR1C1 = "10"
    
' copy to all rows below
    
     Dim lastRow As Long
     With ActiveSheet
        lastRow = .Cells(.Rows.Count, "E").End(xlUp).Row
        Range("E11:E" & lastRow).FormulaR1C1 = "=R[-9]C"
        End With
      'convert to values
      
       ActiveSheet.Range("E11:E" & lastRow).Value = ActiveSheet.Range("E11:E" & lastRow).Value
       
    
' Create Tactical Planning rows for SNP95
    ActiveSheet.UsedRange.AutoFilter Field:=5, Criteria1:="4"
    
  
    
    Application.GoTo Reference:="R1C1"
    ActiveCell.Offset(4, 0).Range("A1:F1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.End(xlDown).Select
    ActiveCell.Offset(6, 0).Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.Offset(0, 4).Range("A1").Select
    ActiveCell.FormulaR1C1 = "7"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Tactical Planning"
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.GoTo Reference:="R1C1"
    
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
        Selection.AutoFilter Field:=5, Criteria1:="7"
    Application.GoTo Reference:="R1C1"
    ActiveCell.Offset(4, 0).Range("A1:F1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.Offset(0, 4).Range("A1").Select
    ActiveCell.FormulaR1C1 = "11"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Value = "In Transit"
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.GoTo Reference:="R1C1"
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Application.GoTo Reference:="R1C1"
    
' Sorts product via cntry via order of Key figures
    Cells.Select
    Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Key2:=Range("d2") _
        , Order2:=xlAscending, Key3:=Range("e2") _
        , Order3:=xlAscending, Header:=xlYes, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom
        
'Adds weeks cover formula in the correct format to SNP

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

Application.GoTo Reference:="R11C8"
ActiveCell.FormulaR1C1 = formula
Range("H11").Select
Selection.Copy
Range("I11").Select
Range(Selection, Selection.End(xlToRight)).Select
ActiveSheet.Paste
Application.CutCopyMode = False


'Adds in relevant stock calculation forumla

    Dim lastColumn As Long
    Application.GoTo Reference:="R1C1"
    Selection.End(xlToRight).Select
    lastColumn = ActiveCell.Column
    
    Application.GoTo Reference:="R12C8"
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(R[-3]C-(R[-3]C[-1]-R[-10]C-R[-9]C-R[-8]C-R[-7]C+R[-6]C+R[-5]C),0)"
    ActiveCell.Select
    Selection.Copy
    Range(Selection, Cells(Selection.Row, lastColumn)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.GoTo Reference:="R1C1"

'Copies stock on hand projection figure
    Application.GoTo Reference:="R1C1"
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
        Selection.AutoFilter Field:=6, Criteria1:= _
    "In Transit"
    
    Dim lastRowNumber As Long
    Application.GoTo Reference:="R1C1"
    lastRowNumber = Selection.End(xlDown).Row
    
    Application.GoTo Reference:="R1C8"
    ActiveCell.Offset(11, 0).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range(Selection, Cells(lastRowNumber, Selection.Column)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
        Selection.AutoFilter Field:=6
    Application.GoTo Reference:="R1C8"
    
'Copies weeks cover figures
    Selection.AutoFilter Field:=6, Criteria1:= _
        "weeks Cover"
    Application.GoTo Reference:="R1C8"
    ActiveCell.Offset(10, 0).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
        Selection.AutoFilter Field:=6
    Application.GoTo Reference:="R1C1"

'Copy-paste all as values only
Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'Adds in relevant stock calculation forumla

    Application.GoTo Reference:="R1C1"
    Selection.End(xlToRight).Select
    lastColumn = ActiveCell.Column
    
    Application.GoTo Reference:="R9C8"
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-1]C="""",RC[-1]+R[3]C-SUM(R[-7]C:R[-4]C)+SUM(R[-3]C:R[-2]C),RC[-1]+R[3]C-SUM(R[-7]C:R[-4]C)+SUM(R[-2]C:R[-1]C))"
    ActiveCell.Select
    Selection.Copy
    Range(Selection, Cells(Selection.Row, lastColumn)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.GoTo Reference:="R1C1"

'Copies stock on hand projection figure
    Application.GoTo Reference:="R1C1"
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
        Selection.AutoFilter Field:=6, Criteria1:= _
    "Stock on hand(proj.)"
    
    Application.GoTo Reference:="R1C1"
    lastRowNumber = Selection.End(xlDown).Row
    
    Application.GoTo Reference:="R1C8"
    ActiveCell.Offset(8, 0).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range(Selection, Cells(lastRowNumber, Selection.Column)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
        Selection.AutoFilter Field:=6
    Application.GoTo Reference:="R1C8"

'Adds in relevant weeks cover calculation forumla
    Application.GoTo Reference:="R11C8"
    ActiveCell.FormulaR1C1 = formula
    Range("H11").Select
    Selection.Copy
    Range("I11").Select
    Range(Selection, Selection.End(xlToRight)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

'Copies weeks cover figure
    Application.GoTo Reference:="R1C1"
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
        Selection.AutoFilter Field:=6, Criteria1:= _
    "weeks Cover"
    
    Application.GoTo Reference:="R1C1"
    lastRowNumber = Selection.End(xlDown).Row
    
'note: this part is very calculation intensive - adding line to temporarily set calcs to manual
Application.Calculation = xlCalculationManual

    Application.GoTo Reference:="R1C8"
    ActiveCell.Offset(10, 0).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range(Selection, Cells(lastRowNumber, Selection.Column)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
        Selection.AutoFilter Field:=6
    Application.GoTo Reference:="R1C8"
    
'This part is where calculations are set back to automatic
Application.Calculation = xlCalculationAutomatic

' Borders
 Application.GoTo Reference:="R2C1"
    ActiveCell.Range("A1:A11").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
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
    
    Application.GoTo Reference:="R2C7"
    ActiveCell.Range("A1:A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.NumberFormat = "#,##0"
    Application.GoTo Reference:="R11C7"
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.NumberFormat = "#,##0.0"
    Application.GoTo Reference:="R2C7"
    ActiveCell.Range("B11").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.NumberFormat = "#,##0"
    Application.GoTo Reference:="R1C1"
    
'Tactical Planning Formatting
    Application.GoTo Reference:="R8C7:R8C34"
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
        Formula1:="="""""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Font
        .Bold = True
    End With
    
    With Selection.FormatConditions(1).Borders(xlLeft)
        .LineStyle = xlContinuous
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 4.99893185216834E-02
        .Weight = xlThin
    End With
    
    With Selection.FormatConditions(1).Borders(xlRight)
        .LineStyle = xlContinuous
        .ThemeColor = 7
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    
    With Selection.FormatConditions(1).Borders(xlTop)
        .LineStyle = xlContinuous
        .ThemeColor = 7
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    
    With Selection.FormatConditions(1).Borders(xlBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 7
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    
    Selection.FormatConditions(1).Interior.Color = 65535

    
    Selection.FormatConditions(1).StopIfTrue = False

'Production formatting
Application.GoTo Reference:="R6C7:R6C34"
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    
    Selection.FormatConditions(1).Font.Bold = True
  
    Selection.FormatConditions(1).Interior.Color = 10092543


    Selection.FormatConditions(1).StopIfTrue = False
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
    
'Copy all formatting and replace zeros
    Application.GoTo Reference:="R2C1"
    ActiveCell.Range("A1:A11").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:="0", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Application.GoTo Reference:="R1C1"
    Application.CutCopyMode = False

'Insert headings and final format
Selection.AutoFilter
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
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
    
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    With Selection.Interior
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.14996795556505
    End With
    
    Selection.Font.Bold = True
    Application.GoTo Reference:="R1C3"
    ActiveCell.FormulaR1C1 = "Cntry"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Loc."
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Ord"
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.GoTo Reference:="C1:C34"
    ActiveCell.Columns("A:AH").EntireColumn.EntireColumn.AutoFit
    Application.GoTo Reference:="R1C1"
    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Application.GoTo Reference:="R1C8"
    ActiveCell.FormulaR1C1 = _
        "=IF(""Wk ""&WEEKNUM(SUBSTITUTE(R[1]C,""."",""/""))=""Wk 53"", ""Wk 1"",""Wk ""&WEEKNUM(SUBSTITUTE(R[1]C,""."",""/"")))"
    ActiveCell.Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.End(xlToRight).Select
    ActiveCell.Offset(-1, 0).Range("A1").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.GoTo Reference:="R1C8:R1C34"
        Selection.Font.Bold = True
    
    With Selection
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
    
    With Selection.Font
        .Name = "Calibri"
        .FontStyle = "Bold"
        .Size = 11
        .ThemeColor = xlThemeColorDark1
        .ThemeFont = xlThemeFontMinor
    End With
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
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
    
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    Selection.Interior.Color = 2162853

    
    Application.GoTo Reference:="R1C1"
    
    Application.GoTo Reference:="C5"
    Selection.EntireColumn.Hidden = True
    Application.GoTo Reference:="R1C1"

    Selection.AutoFilter
    Application.GoTo Reference:="R3C7"
    ActiveWindow.FreezePanes = True
    
    
      
    Sheet1.Range("A1:AH6000").FormatConditions.Delete

    
'Add formatting for Stock on Hand less than zero
    Range("H1:AH6000").Select
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=IF(AND($F1=""Stock on hand(proj.)"",H1<0),TRUE)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    Selection.FormatConditions(1).Font.Color = RGB(156, 0, 6)
    Selection.FormatConditions(1).Interior.Color = RGB(255, 199, 206)

 'Add formatting for Stock on hand below safety stock
    Range("H1:AH6000").Select
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=IF(AND($F1=""Stock on hand(proj.)"",H1<($AM$2*H2),H1>0),TRUE)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Interior.Color = RGB(255, 192, 0)

    
  'Add formatting for Stock on hand over safety stock
    Range("H1:AH6000").Select
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=IF(AND($F1=""Stock on hand(proj.)"",H1>($AM$1*H2)),TRUE)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Interior.Color = RGB(177, 160, 199)

   
   'Add formatting for Stock on hand within safety stock
    Range("H1:AH6000").Select
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=IF(AND($F1=""Stock on hand(proj.)"",H1>=($AM$2*H2),H1<=($AM$1*H2)),TRUE)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    Selection.FormatConditions(1).Interior.Color = RGB(196, 215, 155)
    
    Range("H1:AH6000").Select
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=IF(AND($F1=""Stock on hand(proj.)"",H1=""""),TRUE)"
    
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Interior.ThemeColor = xlThemeColorDark1
    Selection.FormatConditions(1).StopIfTrue = False
    
    
    Range("AL1") = "Overstock - % of Safety Stock(SNP)"
    Range("AL2") = "Understock - % of Safety Stock(SNP)"
    Range("AM1:AM2").Style = "Percent"
    Columns("AL:AL").EntireColumn.AutoFit
    Range("AM1").FormulaR1C1 = "150%"
    Range("AM2").FormulaR1C1 = "75%"
    Range("AL1:AM2").Interior.Color = RGB(149, 179, 215)
    
   'Select A1 and set zoom to 85
    Range("A1").Select
    ActiveWindow.Zoom = 85
    
End Sub








