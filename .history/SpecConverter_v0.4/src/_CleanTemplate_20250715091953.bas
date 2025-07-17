Attribute VB_Name = "Module1"
Sub RemoveNonBWANumberedItems()
    Dim para As Paragraph
    Dim lbl As String
    Dim i   As Long
    
    For i = ActiveDocument.Paragraphs.Count To 1 Step -1
        Set para = ActiveDocument.Paragraphs(i)
        
        ' Safely get the list label (empty string if not a list)
        On Error Resume Next
        lbl = para.Range.ListFormat.ListString
        On Error GoTo 0
        
        ' If it's a list item AND not a BWA- item, delete it
        If Len(lbl) > 0 Then
            If UCase(Left(lbl, 4)) <> "BWA-" Then
                para.Range.Delete
            End If
        End If
    Next i
End Sub

