Attribute VB_Name = "SampleModule"
'@Folder("VBAProject")
Option Explicit

Public Sub IncrementCounter(rc As IRibbonControl)
    Dim r As Range

    Set r = ThisWorkbook.Worksheets(1).Range("$A$2")
    r.Value = r.Value + 1

End Sub

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    ' This subroutine is called when the ribbon is loaded.
    ' You can use this to initialize any necessary variables or state.
End Sub
