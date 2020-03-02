Attribute VB_Name = "Module1"
Sub rb()
cin = InputBox("Enter Reaction")
rct = Split(cin, "->")


rrct = Split(rct(0), "+") 'reactants of reaction

For j = LBound(rrct) To UBound(rrct)
MsgBox (j)
Next j
'rct (0) + "+" + rct(1)

prct = Split(rct(1), "+") 'products of reaction

For j = LBound(prct) To UBound(prct)
MsgBox (prct(j))
Next j

End Sub
