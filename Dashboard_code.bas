Attribute VB_Name = "Module2"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    With selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
        .Transparency = 0
    End With
    With selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    
End Sub
Sub pdst()
Dim pdsni(1 To 36) As Variant

'For iter = 0 To 5
    For c = 1 To 36
        pdsni(c) = Format(Worksheets("PDSni").Range("L10").Offset(c, 0).Value, "0.###")
    Next c
    
temp = ""
For c = 1 To 36
    temp = temp & "-" & pdsni(c)
Next c
MsgBox temp
End Sub
