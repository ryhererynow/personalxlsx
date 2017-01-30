Attribute VB_Name = "Module3"
Sub Text2C()
Attribute Text2C.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Text2C Macro
'

'
    Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar _
        :=".", FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Columns("B:B").TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar _
        :=".", FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Columns("C:C").TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar _
        :=".", FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Columns("D:D").TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar _
        :=".", FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Columns("E:E").TextToColumns Destination:=Range("E1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar _
        :=".", FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Columns("F:F").TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar _
        :=".", FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
End Sub
