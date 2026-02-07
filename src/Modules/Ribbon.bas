Option Explicit

Public gRibbon As IRibbonUI

' Called by Excel when loading the add-in
Public Sub OnLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
End Sub

' Callback for button click
Public Sub OnRunAutomation(control As IRibbonControl)
    RunAutomation
End Sub
