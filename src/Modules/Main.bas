Option Explicit

Public Sub RunAutomation()
    Dim response As VbMsgBoxResult

    response = MsgBox( _
        "This will process the selected workbook and generate output." & vbCrLf & _
        "Do you want to continue?", _
        vbQuestion + vbOKCancel, _
        "Confirm")

    If response <> vbOK Then Exit Sub

    ' Show range picker form
    RangePicker.Show
End Sub
