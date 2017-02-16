VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xlsToPpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private powerPointApp As Object
Private currentPresentation As Object
Public pasteDictionary As Object
Public Enum powerPointPastes

    ppPasteBitmap = 1
    ppPasteDefault = 0
    ppPasteEnhancedMetafile = 2
    ppPasteGIF = 4
    ppPasteHTML = 8
    ppPasteJPG = 5
    ppPasteMetafilePicture = 3
    ppPasteOLEObject = 10
    ppPastePNG = 6
    ppPasteRTF = 9
    ppPasteShape = 11
    ppPasteText = 7

End Enum


Private Sub Class_Initialize()



Set pasteDictionary = CreateObject("Scripting.Dictionary")

pasteDictionary.Add "ppPasteBitmap", 1
pasteDictionary.Add "ppPasteDefault", 0
pasteDictionary.Add "ppPasteEnhancedMetafile", 2
pasteDictionary.Add "ppPasteGIF", 4
pasteDictionary.Add "ppPasteHTML", 8
pasteDictionary.Add "ppPasteJPG", 5
pasteDictionary.Add "ppPasteMetafilePicture", 3
pasteDictionary.Add "ppPasteOLEObject", 10
pasteDictionary.Add "ppPastePNG", 6
pasteDictionary.Add "ppPasteRTF", 9
pasteDictionary.Add "ppPasteShape", 11
pasteDictionary.Add "ppPasteText", 7


End Sub


Public Function OpenOrCreatePowerPoint(Optional ByVal presentationDirectory As String) As Object

Dim powerPointApp As Object
Dim slideShow As Object
Set powerPointApp = BingOrCreatePPT()

If presentationDirectory = "" Then

    
    If currentPresentation Is Nothing Then
    
        Set slideShow = powerPointApp.Presentations.Add
    
    Else
    
        Set slideShow = currentPresentation
    
    End If
Else



    On Error Resume Next
    
        Set slideShow = powerPointApp.Presentations.Open(presentationDirectory)
        
        If Err.Number <> 0 Then
        
            errorString = "Failed to open directory: " & presentationDirectory & _
            "the file may have been moved or deleted, alternatively you may lack privlages to open it, "
            errors.AddError errorString, CriticalError
            
        End If
        
    On Error GoTo 0
    
End If

Set OpenOrCreatePowerPoint = slideShow
Set currentPresentation = slideShow

End Function

Private Function BingOrCreatePPT() As Object

Dim oPPT As Object

On Error Resume Next

    Set oPPT = GetObject(, "PowerPoint.Application")

    If Err.Number <> 0 Then

       Set oPPT = CreateObject("PowerPoint.Application")
        
    End If
On Error GoTo 0

Set BingOrCreatePPT = oPPT
Set powerPointApp = oPPT

End Function


Public Function AddSlide(Optional ByVal before As Long, Optional ByVal templateID As Integer) As Object

    If currentPresentation Is Nothing Then
    
        Set currentPresentation = OpenOrCreatePowerPoint
        
    End If
    
    If templateID = 0 Then
    
        templateID = 12
        
    End If
    
    With currentPresentation.slides
    
        If before > .Count Then
        
            before = .Count + 1
            
        End If
        If before = 0 Then
        
            Set AddSlide = .Add(.Count + 1, templateID)
        
        Else
        
            Set AddSlide = .Add(before, templateID)
            
        End If
    End With
End Function
Public Function DeleteSlide(slideID As Long)
Dim errors As New errorHandler
If currentPresentation Is Nothing Then

    errors.AddError "You are trying to delete a slide while a presentation is not set", Warning
    Set errors = Nothing
    
Else

    currentPresentation.slides(slideID).Delete
    
End If

End Function

Public Function PasteChartToSlide(ByRef excelChart As Chart, ByVal slideID As Long, ByVal top As Long, ByVal left As Long _
                            , ByVal height As Long, ByVal width As Long, Optional ByVal pasteType As powerPointPastes)

Dim currentChart As Variant
Dim errors As New errorHandler
Dim pasteINT As Integer
If currentPresentation Is Nothing Then
    
    Set currentPresentation = OpenOrCreatePowerPoint
        
End If



If pasteType = 0 Then

    pasteType = ppPastePNG
    
End If


If currentPresentation.slides.Count = 0 Then

    AddSlide
    
End If
    
If currentPresentation.slides.Count < slideID Then

    errors.AddError "Slide" & slideID & " does not exists", StandardError

Else

    On Error Resume Next
    
        If pasteDictionary(pasteType) Is Empty Then
        
            pasteINT = pasteType
            
        Else
        
            pasteINT = pasteDictionary(pasteType)
            
        End If
        excelChart.ChartArea.Copy
        
        Set currentChart = currentPresentation.slides(slideID).Shapes.PasteSpecial(pasteINT)
        
        With currentChart
        
            .LockAspectRatio = False
            .top = top
            .left = left
            .height = height
            .width = width
        
        End With
        
        
        If Err.Number <> 0 Then
        
            errors.AddError "failed to paste chart: " & currentChart.Name & ", ", StandardError
            
        End If

    On Error GoTo 0
End If

End Function
                            
Public Function PasteRangeToSlide(ByRef excelRange As Range, ByVal slideID As Long, ByVal top As Long, ByVal left As Long _
                            , ByVal height As Long, ByVal width As Long, Optional ByVal pasteType As powerPointPastes)
                            
Dim currentRange As Range
Dim errors As New errorHandler
Dim pasteINT As Integer
If currentPresentation Is Nothing Then
    
    Set currentPresentation = OpenOrCreatePowerPoint
        
End If



If pasteType = 0 Then

    pasteType = ppPasteMetafilePicture
    
End If



If currentPresentation.slides.Count = 0 Then

    AddSlide
    
End If
    
If currentPresentation.slides.Count < slideID Then

    errors.AddError "Slide" & slideID & " does not exists", StandardError

Else

    On Error Resume Next
        If pasteDictionary(pasteType) Is Empty Then
        
            pasteINT = pasteType
            
        Else
        
            pasteINT = pasteDictionary(pasteType)
            
        End If
        
        Select Case pasteINT
        
            Case 4, 5, 6
                pasteINT = 3
                errors.AddError excelRange.Parent.Name & "-" & excelRange.Address & _
                " PNG, GIF, JPEG conversions are not supported by Excel converted to metafile picture", Warning
        
        End Select
        
        excelRange.Copy
        
                Set currentChart = currentPresentation.slides(slideID).Shapes.PasteSpecial(pasteINT)
        
        With currentChart
        
            .LockAspectRatio = False
            .top = top
            .left = left
            .height = height
            .width = width
        
        End With
        
        
        If Err.Number <> 0 Then
        
            errors.AddError "failed to paste chart: " & currentChart.Name & ", ", StandardError
            
        End If

    On Error GoTo 0
End If

End Function

Public Function FillInText(ByVal fillString As String, ByVal slideID As Long, ByVal textObjectName As String, Optional ByVal top As Long, _
                            Optional ByVal left As Long, Optional ByVal height As Long, Optional ByVal width As Long)
                            
Dim currentTextBox As Variant
                            
If currentPresentation Is Nothing Then
    
    Set currentPresentation = OpenOrCreatePowerPoint
        
End If

If currentPresentation.slides.Count = 0 Then

    AddSlide
        
End If

If currentPresentation.slides.Count < slideID Then

    errors.AddError "Slide" & slideID & " does not exists", StandardError

Else
    On Error Resume Next
    
        Set currentTextBox = currentPresentation.slides(slideID).Shapes(textObjectName)
        
        If Err.Number <> 0 Then
        
            Set currentTextBox = currentPresentation.slides(slideID).Shapes.AddTextbox(msoTextOrientationHorizontal, left, top, width, height)
            With currentTextBox
                
                .left = 1
                .top = 1
                .width = 100
                .height = 100
                
            End With
        End If
           
    
        currentTextBox.TextFrame.TextRange.Text = fillString
        
        If IsMissing(top) Then
        
            top = currentTextBox.top
        
        End If
        
        If IsMissing(left) Then
        
            left = currentTextBox.left
            
        End If
        
        
        If IsMissing(height) Then
        
            height = currentTextBox.height
            
        End If
        
        If IsMissing(width) Then
        
            width = currentTextBox.width
            
        End If
        
        
        With currentTextBox
        
            .LockAspectRatio = False
            .top = top
            .left = left
            .height = height
            .width = width
            .Name = textObjectName
            
        End With
        
        If Err.Number <> 0 Then
        
            errors.AddError "Did not find item: " & textObjectName & ", inserted a new textbox ", Warning
            
        End If
    
    On Error GoTo 0
End If
End Function

