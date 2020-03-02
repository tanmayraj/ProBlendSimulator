Attribute VB_Name = "Module3"
Sub tearblocktrial()

Set bmo = Worksheets("BallMill").Range("L8")
Set pdsrc = Worksheets("PDSTearSource").Range("L8")
Set pdsout = Worksheets("PDSxc").Range("L393")
Set pdsni = Worksheets("PDSni").Range("L8")
Set rr = Worksheets("PDSTearBlock").Range("RecycleRatio")

Itbsrc = Worksheets("PDSTearBlock").Range("k:k").Find("PDS Tear Source").Address
Set Itbsrc = Worksheets("PDSTearBlock").Range(Itbsrc).Offset(1, 3)

Itbni = Worksheets("PDSTearBlock").Range("k:k").Find("PDS Net Input").Address
Set Itbni = Worksheets("PDSTearBlock").Range(Itbni).Offset(1, 3)

itbts = Worksheets("PDSTearBlock").Range("k:k").Find("PDS Tear Sink").Address
Set itbts = Worksheets("PDSTearBlock").Range(itbts).Offset(1, 3)

'------------------------------------------------------------------------
'Calculate Iteration
Set pds_source = Worksheets("PDSTearBlock").Range("B3").Offset(2, 0)
Set pds_out = Worksheets("PDSTearBlock").Range("C3").Offset(2, 0)
Set pds_sink = Worksheets("PDSTearBlock").Range("D3").Offset(2, 0)
Set ino = Worksheets("PDSTearBlock").Range("A3").Offset(2, 0)

ino.Offset(0, 0) = 0
pds_source.Offset(0, 0) = 0
pds_out.Offset(0, 0) = bmo.Offset(-5, 0) + 6
pds_sink.Offset(0, 0) = pds_out.Offset(0, 0) * rr
c = 0
Do While pds_sink.Offset(c, 0) - pds_source.Offset(c, 0) > 0.000001
    c = c + 1
    
    ino.Offset(c, 0) = ino.Offset(c - 1, 0) + 1
    pds_source.Offset(c, 0) = pds_sink.Offset(c - 1, 0)
    pds_out.Offset(c, 0) = pds_out + pds_source.Offset(c, 0)
    pds_sink.Offset(c, 0) = pds_out.Offset(c, 0) * rr
    
    'MsgBox c
    'MsgBox (pds_sink.Offset(c, 0) - pds_source.Offset(c, 0))
Loop
'MsgBox ("Iterations are: c")
Simulator.TextBox1.Value = c
'========================================================================
For iter = 0 To c


            
    If iter = 0 Then
        For j = 2 To 59
            If j <> 38 And j <> 39 And j <> 40 And j <> 55 And j <> 56 And j <> 57 Then
                Itbsrc.Offset(j, 0) = 0                                                     'initializing tear source as empty
            End If
        Next j
            
        For j = 2 To 59
            If j <> 38 And j <> 39 And j <> 40 And j <> 55 And j <> 56 And j <> 57 Then
                pdsrc.Offset(j, 0) = Itbsrc.Offset(j, 0)                                    'copying inital tear source value into stream
            End If
        Next j
                                                                                            'now PDSni stream is automatically calculated from sheet
        For j = 2 To 59
            If j <> 38 And j <> 39 And j <> 40 And j <> 55 And j <> 56 And j <> 57 Then
                Itbni.Offset(j, 0) = bmo.Offset(j, 0) + Itbsrc.Offset(j, 0)                 'PDSni is recalculated and entered in Tear Block
            End If
        Next j
                                                                                            'PDSxc stream is calculated automatically from sheet
        For j = 2 To 59
            If j <> 38 And j <> 39 And j <> 40 And j <> 55 And j <> 56 And j <> 57 Then
                itbts.Offset(j, 0) = rr * pdsout.Offset(j, 0)                               'PDS Tear Sink value is calculated in Tear Block
            End If
        Next j
            
            
    Else
    
        For j = 2 To 59
            If j <> 38 And j <> 39 And j <> 40 And j <> 55 And j <> 56 And j <> 57 Then
                Itbsrc.Offset(j, iter) = itbts.Offset(j, iter - 1)                                 'Previous iteration Tear Sink is now the Tear Source
            End If
        Next j
        
        For j = 2 To 59
            If j <> 38 And j <> 39 And j <> 40 And j <> 55 And j <> 56 And j <> 57 Then
                pdsrc.Offset(j, 0) = Itbsrc.Offset(j, iter)                                 'The new Tear Source value is entered into the stream
            End If
        Next j
                                                                                            'PDSni stream is automatically updated with worksheet formula and so is the PDSxc output value
        For j = 2 To 59
            If j <> 38 And j <> 39 And j <> 40 And j <> 55 And j <> 56 And j <> 57 Then
                Itbni.Offset(j, iter) = bmo.Offset(j, 0) + Itbsrc.Offset(j, iter)           'PDSni is recalculated into Tear Block
            End If
        Next j
        
        For j = 2 To 59
            If j <> 38 And j <> 39 And j <> 40 And j <> 55 And j <> 56 And j <> 57 Then
                itbts.Offset(j, iter) = rr * pdsout.Offset(j, 0)                                   'New Tear Sink value is calculated
            End If
        Next j
    End If
    

'If iter Mod 5 = 0 Then
Simulator.Progress.Width = (iter / c) * 193
Simulator.Label2.Caption = Format((iter / c) * 100, "0") & "%"
DoEvents
If iter = c Then
MsgBox "Complete"
Exit Sub

End If
'End If
'MsgBox iter
Next iter                                                                                   'iter + 1

End Sub

