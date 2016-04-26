Sub changefont()
'This macro changes fonts of the title text

Dim osld As Slide, oshp As Shape

	For Each osld In ActivePresentation.Slides
		For Each oshp In osld.Shapes
			If oshp.Type = msoPlaceholder Then
'Title text change values as required
				If oshp.PlaceholderFormat.Type = 1 Or oshp.PlaceholderFormat.Type = 3 Then
					If oshp.HasTextFrame Then
						If oshp.TextFrame.HasText Then
							With oshp.TextFrame.TextRange.Font
								.Name = "Arial"
								.Size = 36
								.Color.RGB = RGB(0, 0, 255)
								.Bold = msoFalse
								.Italic = msoFalse
								.Shadow = False
							End With
						End If
					End If
				End If
				If oshp.PlaceholderFormat.Type = 2 Or oshp.PlaceholderFormat.Type = 7 Then
					If oshp.HasTextFrame Then
						If oshp.TextFrame.HasText Then
							'Body text change values as required
							With oshp.TextFrame.TextRange.Font
								.Name = "Arial"
								.Size = 24
								.Color.RGB = RGB(255, 0, 0)
								.Bold = msoFalse
								.Italic = msoFalse
								.Shadow = False
							End With
						End If
					End If
				End If
			End If
		Next oshp
	Next osld
	
End Sub

