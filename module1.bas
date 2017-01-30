Attribute VB_Name = "Module1"
Sub Count_into_Sum()

Dim mytbl As PivotTable
Set mytbl = ActiveSheet.PivotTables(1)

Dim fld As PivotField
For Each fld In mytbl.DataFields
    If fld.Function = xlCount Then
    fld.Function = xlSum
End If
fld.NumberFormat = "#,##0"
Next

End Sub

Sub AsianDates()
'
' AsianDates Macro
'

Selection.NumberFormat = "yyyy-mm-dd;@"


End Sub

