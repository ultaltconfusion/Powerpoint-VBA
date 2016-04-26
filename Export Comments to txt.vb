Sub ExportComments()

'This takes all of the comments out of the active presentation
'And outputs it into a .txt file within the same directory
'As the active presentation

    Dim iFile As Integer      'File handle for output
    Dim PathSep As String
    iFile = FreeFile          'Get a free file number

    'If Mac Then
    ' PathSep = ":"
    'Else
    PathSep = "\"
    'End If

    Dim oSl As Slide
    Dim oSlides As Slides
    Dim oCom As Comment
    Dim oRep As Comment
    Dim sText As String
    Dim sFilename As String

    Set oSlides = ActivePresentation.Slides
    
    Open ActivePresentation.Path & PathSep & "AllComments.txt" For Output As iFile
        
    For Each oSl In oSlides
        If oSl.Comments.Count > 0 Then
    
            sText = sText & "Slide: " & oSl.SlideIndex & vbCrLf
            sText = sText & "======================================" & vbCrLf
                
                For Each oCom In oSl.Comments
                
                    sText = sText & oCom.Author & vbCrLf
                    sText = sText & oCom.DateTime & vbCrLf
                    sText = sText & oCom.Text & vbCrLf
                    sText = sText & "--------------" & vbCrLf
                    
                    If oCom.Replies.Count > 0 Then
                        For Each oRep In oCom.Replies
                            sText = sText & "REPLY TO ABOVE:" & vbCrLf
                            sText = sText & oRep.Author & vbCrLf
                            sText = sText & oRep.DateTime & vbCrLf
                            sText = sText & oRep.Text & vbCrLf
                            sText = sText & "--------------" & vbCrLf
                        Next oRep
                    End If
                
                sText = sText & "**************" & vbCrLf
                Next oCom
                
                sText = sText & vbCrLf
        End If
    
    Next oSl

    Print #iFile, sText
    Close #iFile

End Sub
