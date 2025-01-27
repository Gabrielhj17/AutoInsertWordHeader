Sub InsertHeaderInfo()
    Dim headerRange As Range
    Dim doc As Document

    ' Set the document variable to the active document
    Set doc = ActiveDocument
    
    ' Create a reference to the header range of the first page
    Set headerRange = doc.Sections(1).Headers(wdHeaderFooterPrimary).Range
    
    ' Clear the existing content of the header
    headerRange.Delete
    
    ' Insert the name and date in the header with the desired formatting
    headerRange.ParagraphFormat.Alignment = wdAlignParagraphLeft ' Align text to the left
    headerRange.Text = "YOUR NAME"

    ' Insert the date with a right-aligned tab character to align it to the right
    headerRange.InsertAfter Format(Now, "                                                                                                           dd/mm/yyyy")
End Sub
