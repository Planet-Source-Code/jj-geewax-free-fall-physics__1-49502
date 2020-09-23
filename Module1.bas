Attribute VB_Name = "Module1"
Option Explicit

Public Function FindInitialHeight(GroundT As Double, _
                                  Vel As Double, _
                                  Acc As Double) As Double

    FindInitialHeight = (Vel * GroundT) + (0.5 * Acc * GroundT * GroundT)

End Function

Public Function FindInitialVelocity(InitialHeight As Double, _
                                    GroundT As Double, _
                                    Acc As Double) As Double

    If GroundT = 0 Then
        Exit Function
    End If
    FindInitialVelocity = (InitialHeight / GroundT) - (0.5 * Acc * GroundT)

End Function

Public Function FindTimeHitsGround(Acc As Double, _
                                   Vi As Double, _
                                   Hi As Double) As Double

    FindTimeHitsGround = SolveQuadratic(Acc / 2, Vi, -1 * Hi)

End Function

Public Function FindVelocity(Acc As Double, _
                             Time As Double) As Double

    FindVelocity = Acc * Time

End Function

Public Function SolveQuadratic(ByVal A As Double, _
                               ByVal B As Double, _
                               ByVal C As Double) As Double

  '<:-):WARNING: 'ByVal ' inserted for Parameters 'A As Double, B As Double, C As Double'
  
  Dim First  As Double
  Dim Second As Double

    '<:-):UPDATED: Multiple Dim line separated
    First = -B + (Sqr((B * B) - (4 * A * C)))
    If Not A = 0 Then
        First = First / (2 * A)
    End If
    If First > 0 Then
        SolveQuadratic = First
        Exit Function
    End If
    '======================
    Second = -B - (Sqr((B * B) - (4 * A * C)))
    If Not A = 0 Then
        Second = Second / (2 * A)
    End If
    If Second > 0 Then
        SolveQuadratic = Second
        Exit Function
     Else
        SolveQuadratic = 0
    End If

End Function

Public Sub SaveText(txtSave As RichTextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave.Text
    Open Path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub
Public Sub SaveListBox(Directory As String, TheList As ListBox)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&
    Close #1
End Sub
