# Find And Highlight Duplicate Paragraphs In ResearchArticle or Thesis

Suppose you have a large research article or thesis in word document format which may have hundreds of pages. As part of editing, you want to check if there are duplicate paragraphs and then highlight to make them outstanding, so that you can deal with the duplicate sentences.

To do the same, we use a VBA (Microsoft Visual Basic for Applications window) code.
For activating VBA code in Microsoft word. Open your desired word file and press Alt+F11 in your keyboard. Thus will open the VBA window.

- Click Insert > Module. This will open another window.
- Copy and paste below code into the opened blank module
- Now, press F5 key to run this code, all the duplicate sentences are highlighted at once, the first displayed duplicate paragraphs are highlighted with green color, and other duplicates are highlighted with yellow color.

```
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
```


