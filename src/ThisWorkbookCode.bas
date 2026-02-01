Option Explicit

Private Sub Workbook_Open()
    ' Keep it simple & robust
    On Error Resume Next
    SetAddInMetadata
    On Error GoTo 0
End Sub

Private Sub SetAddInMetadata()
    ' These usually drive the “nice” name/description shown in the Add-ins dialog.
    SetBuiltInDocProp "Title", "CSV Exporter v1.3.0"
    SetBuiltInDocProp "Comments", _
        "CSV Exporter – Enhanced fork by Shep Bivens" & vbCrLf & _
        "Based on original work by Brian Skinn"
End Sub

Private Sub SetBuiltInDocProp(ByVal propName As String, ByVal propValue As String)
    Dim p As Object
    Set p = ThisWorkbook.BuiltinDocumentProperties(propName)
    p.Value = propValue
End Sub
