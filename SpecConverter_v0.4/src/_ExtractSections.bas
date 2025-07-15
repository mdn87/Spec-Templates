Attribute VB_Name = "Module1"
Sub DetectSections_WithSpecLevels()
    Dim p        As Paragraph
    Dim listStr  As String, title As String
    Dim wordLvl  As Long, specLevel As Long
    Dim leftInd  As Single, firstInd As Single
    Dim parts()  As String
    
    For Each p In ActiveDocument.Paragraphs
        ' 1) grab Word numbering (empty if manual)
        On Error Resume Next
        listStr = p.Range.ListFormat.ListString
        On Error GoTo 0
        
        If listStr <> "" Then
            ' capture Word’s level (for reference, if you need it)
            wordLvl = p.Range.ListFormat.ListLevelNumber
            
            ' compute your spec level = number of segments - 1
            parts = Split(listStr, ".")
            specLevel = UBound(parts)
            
            ' grab the full run-text, since Word omits the number
            title = Trim(p.Range.Text)
            
        Else
            ' you can keep your RegExp fallback here if needed…
            GoTo NextPara
        End If
        
        ' indent info
        leftInd = p.LeftIndent
        firstInd = p.FirstLineIndent
        
        ' output (or write to XML) using specLevel, not wordLvl
        Debug.Print "SECTION: " & listStr & " ? " & title
        Debug.Print "   specLevel=" & specLevel & _
                    "   Indent=" & leftInd & "pt" & _
                    "   FirstIndent=" & firstInd & "pt"
NextPara:
    Next p
End Sub


