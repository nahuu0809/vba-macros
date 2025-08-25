Sub CreateWordDocumentsInBulk()
    Dim i As Integer
    Dim doc As Document
    Dim docName As String
    Dim folderPath As String
    
    ' Set the folder path where you want to save the documents
    folderPath = "C:\Your\Desired\Path\"
    
    ' Loop to create 20 documents
    For i = 1 To 20
        ' Create a new document
        Set doc = Documents.Add
        
        ' Add some content to the document (optional)
        doc.Content.Text = "This is document number " & i
        
        ' Set the document name
        docName = folderPath & "Document_" & i & ".docx"
        
        ' Save the document
        doc.SaveAs2 FileName:=docName, FileFormat:=wdFormatXMLDocument
        
        ' Close the document
        doc.Close SaveChanges:=False
    Next i
    
    ' Inform the user that the process is complete
    MsgBox "20 documents have been created successfully.", vbInformation
End Sub
