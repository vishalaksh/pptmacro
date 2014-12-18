Attribute VB_Name = "Module1"
Sub use_regex()
    Dim regX As Object
    Dim osld As Slide
    Dim oshp As Shape
    Dim strInput As String
    Dim b_found As Boolean
    Dim iRow As Integer
    Dim iCol As Integer

    Set regX = CreateObject("vbscript.regexp")
    With regX
        .Global = True
        .Pattern = "(\d)"
    End With
    For Each osld In ActivePresentation.Slides
        For Each oshp In osld.Shapes
            If oshp.HasTable Then
                For iRow = 1 To oshp.Table.Rows.Count
                    For iCol = 1 To oshp.Table.Columns.Count
                        strInput = oshp.Table.Cell(iRow, iCol).Shape.TextFrame.TextRange.Text
                        b_found = regX.Test(strInput)
                        If b_found = True Then
                        
                             Set myMatches = regX.Execute(strInput)
                             For Each myMatch In myMatches
                            oshp.Table.Cell(iRow, iCol).Shape.TextFrame.TextRange.Characters(myMatch.FirstIndex + 1, myMatch.Length).Characters.Font.Name = "Times New Roman"
                            Next
                        End If
                    Next iCol
                Next iRow
            Else
                If oshp.HasTextFrame Then
                    If oshp.TextFrame.HasText Then
                        strInput = oshp.TextFrame.TextRange.Text
                        b_found = regX.Test(strInput)
                        If b_found = True Then
                         
                            Set myMatches = regX.Execute(strInput)
                                For Each myMatch In myMatches
                                    oshp.TextFrame.TextRange.Characters(myMatch.FirstIndex + 1, myMatch.Length).Characters.Font.Name = "Times New Roman"
                                Next
                          
                        End If
                    End If
                End If
            End If
        Next oshp
    Next osld
    Set regX = Nothing
End Sub

