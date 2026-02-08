Option Explicit

Private Sub UserForm_Initialize()
    lblInstructions.Caption = "Select a range to process:"
End Sub

Private Sub btnPick_Click()
    On Error Resume Next

    Dim rng As Range
    Set rng = Application.InputBox( _
        "Select a range", _
        "Range Selection", _
        Type:=8)

    On Error GoTo 0

    If Not rng Is Nothing Then
        Set SelectedRange = rng
        txtRange.Text = rng.Address(External:=True)
    End If
End Sub

Private Sub btnOK_Click()
    If SelectedRange Is Nothing Then
        MsgBox "Please select a range first.", vbExclamation
        Exit Sub
    End If

    Me.Hide

    ' Placeholder: call real logic here
    ProcessSelectedRange SelectedRange
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
