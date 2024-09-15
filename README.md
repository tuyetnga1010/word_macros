# word_macros
# Resize all pictures
Sub Resizepic()

    Dim shp As InlineShape
    Dim targetWidth As Single
    Dim targetHeight As Single

    targetWidth = 100
    targetHeight = 100

    For Each shp In ActiveDocument.InlineShapes
        With shp
            .LockAspectRatio = msoFalse
            .Width = targetWidth
            .Height = targetHeight
        End With
    Next shp

End Sub
