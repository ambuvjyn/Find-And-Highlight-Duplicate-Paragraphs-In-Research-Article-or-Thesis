Sub HighlightDuplicateParagraphs()
    Dim i As Integer, j As Integer
    Dim Original As Paragraph, Duplicate As Paragraph
    Dim OriginalText As String, DuplicateText As String
    
    For i = 1 To ActiveDocument.Paragraphs.Count - 1
        Set Original = ActiveDocument.Paragraphs(i)
        OriginalText = Trim(Original.Range.Text)
        If Len(OriginalText) > 0 Then
            For j = i + 1 To ActiveDocument.Paragraphs.Count
                Set Duplicate = ActiveDocument.Paragraphs(j)
                DuplicateText = Trim(Duplicate.Range.Text)
                If Len(DuplicateText) > 0 And OriginalText = DuplicateText Then
                    Original.Range.HighlightColorIndex = wdGreen
                    Duplicate.Range.HighlightColorIndex = wdYellow
                    Exit For
                End If
            Next j
        End If
    Next i
End Sub
