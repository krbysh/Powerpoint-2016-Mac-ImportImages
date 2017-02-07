Sub importImages()

Dim folPath As String
Dim strFileSpec As String
Dim oSld As Slide
Dim oPic As Shape

Dim fName As String
Dim fNull As String
Dim fNum As Integer
Dim i As Integer
Dim fArray() As String
Dim fAccGranted As Boolean

' Edit these to suit:
' folPath is also used by AppleScript as aFol
' it isn't necessary to add the last "slash"

folPath = "/Users/krbysh/Documents/98_Cap"


' Call AppleScript
fNull = AppleScriptTask("importPngImages.scpt", "myFolderName", folPath)
fNum = AppleScriptTask("importPngImages.scpt", "myFileNum", vbNullString)

ReDim fArray(0 To fNum - 1)

For i = 1 To fNum
    fName = AppleScriptTask("importPngImages.scpt", "myFileName", i)
    fArray(i - 1) = folPath & "/" & fName
Next i

fAccGranted = GrantAccessToMultipleFiles(fArray)

For i = 1 To fNum
    Set oSld = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutText)
    Set oPic = oSld.Shapes.AddPicture(fileName:=fArray(i - 1), _
    LinkToFile:=msoFalse, _
    SaveWithDocument:=msoTrue, _
    Left:=27, _
    Top:=78)

    With oPic
      .LockAspectRatio = msoCTrue
      If .Width >= 888 Or .Height >= 290 Then
        If .Width > 888 Then
           .Width = 888
        Else
           .Height = 290
        End If
      End If
    End With
Next i

End Sub

