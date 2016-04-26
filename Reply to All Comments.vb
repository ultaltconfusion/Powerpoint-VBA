Sub addcommentCBP()

    Dim oSl As Slide
    Dim oSlides As Slides
    Dim oCom As Comment

    Set oSlides = ActivePresentation.Slides
            
    For Each oSl In oSlides
        If oSl.Comments.Count > 0 Then
                For Each oCom In oSl.Comments
                    If oCom.Replies.Count > 0 Then
                        Set oCom.Replies = oCom.Replies.Add(oCom.Top, oCom.Left, "CBPartners", "CBP", "")
                    End If
                Next oCom
        End If
    Next oSl

End Sub