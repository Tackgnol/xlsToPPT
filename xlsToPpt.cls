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

Public Enum powerPointAligement

    ppAlignCenter = 2
    ppAlignDistribute = 5
    ppAlignJustify = 4
    ppAlignJustifyLow = 7
    ppAlignLeft = 1
    ppAlignmentMixed = -2
    ppAlignRight = 3
    ppAlignThaiDistribute = 6
    
End Enum

Public Enum TableBorders

    ppBorderTop = 1
    ppBorderRight = 4
    ppBorderLeft = 2
    ppBorderBottom = 3

End Enum
Property Get presentation()

    Set presentation = currentPresentation

End Property


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


Public Function OpenOrCreatePowerPoint(Optional ByVal presentationDirectory As String, Optional ByVal force4x3 As Boolean) As Object

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

If force4x3 Then

    slideShow.PageSetup.SlideSize = 1

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
    
    With currentPresentation.Slides
    
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

Public Function CopySlide(ByVal sourceSlideID As Long) As Long

    Dim currentSlide As Object
    
    If currentPresentation Is Nothing Then
    
        Set currentPresentation = OpenOrCreatePowerPoint
        
    End If
    
    Set currentSlide = currentPresentation.Slides(sourceSlideID).Duplicate
    
    CopySlide = currentSlide.SlideNumber

End Function

Public Function DeleteSlide(slideID As Long)
Dim errors As New errorHandler
If currentPresentation Is Nothing Then

    errors.AddError "You are trying to delete a slide while a presentation is not set", Warning
    Set errors = Nothing
    
Else

    currentPresentation.Slides(slideID).Delete
    
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


If currentPresentation.Slides.Count = 0 Then

    AddSlide
    
End If
    
If currentPresentation.Slides.Count < slideID Then

    errors.AddError "Slide" & slideID & " does not exists", StandardError

Else

    On Error Resume Next
    
        If pasteDictionary(pasteType) Is Empty Then
        
            pasteINT = pasteType
            
        Else
        
            pasteINT = pasteDictionary(pasteType)
            
        End If
        excelChart.ChartArea.Copy
        
        Set currentChart = currentPresentation.Slides(slideID).Shapes.PasteSpecial(pasteINT)
        
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



If currentPresentation.Slides.Count = 0 Then

    AddSlide
    
End If
    
If currentPresentation.Slides.Count < slideID Then

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
        
                Set currentChart = currentPresentation.Slides(slideID).Shapes.PasteSpecial(pasteINT)
        
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

If currentPresentation.Slides.Count = 0 Then

    AddSlide
        
End If

If currentPresentation.Slides.Count < slideID Then

    errors.AddError "Slide" & slideID & " does not exists", StandardError

Else
    On Error Resume Next
    
        Set currentTextBox = currentPresentation.Slides(slideID).Shapes(textObjectName)
        
        If Err.Number <> 0 Then
        
            Set currentTextBox = currentPresentation.Slides(slideID).Shapes.AddTextBox(msoTextOrientationHorizontal, left, top, width, height)
            With currentTextBox
                
                .left = 1
                .top = 1
                .width = 100
                .height = 100
                
            End With
        End If
           
    
        currentTextBox.TextFrame.TextRange.text = fillString
        
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
Public Function AddTable(ByVal Rows As Integer, ByVal Columns As Integer, _
                               ByVal slideID As Long, ByVal top As Long, ByVal left As Long, _
                               ByVal height As Long, ByVal width As Long) As Object

slideID = Builder(slideID)

Set AddTable = currentPresentation.Slides(slideID).Shapes.AddTable(Rows, Columns, left, top, width, height)

End Function


Public Function FormatTable(ByRef Table As Object, ByVal backgroundColor As Long, _
                             ByVal BorderWidth As Integer, ByVal BorderColor As Long, _
                             Optional outBorderWidth As Integer, Optional outBorderColor As Long)
                             
For Each row In Table.Table.Rows

    For Each Cell In row.Cells
        With Cell
        
            .Borders(3).weight = BorderWidth
            .Borders(2).weight = BorderWidth
            .Borders(4).weight = BorderWidth
            .Borders(1).weight = BorderWidth
            .Borders(3).ForeColor.RGB = BorderColor
            .Borders(2).ForeColor.RGB = BorderColor
            .Borders(4).ForeColor.RGB = BorderColor
            .Borders(1).ForeColor.RGB = BorderColor
            .Shape.Fill.BackColor.RGB = backgroundColor
            .Shape.Fill.ForeColor.RGB = backgroundColor
        End With
    
    Next Cell
Next row

If Not outBorderWidth = 0 Or outBorderColor = 0 Then
With Table.Table
    For Each Cell In .Rows(1).Cells
    
        Cell.Borders(1).weight = outBorderWidth
        Cell.Borders(1).ForeColor.RGB = outBorderColor
    
    Next Cell
    
    For Each Cell In .Columns(1).Cells
    
        Cell.Borders(2).weight = outBorderWidth
        Cell.Borders(2).ForeColor.RGB = outBorderColor
    
    Next Cell
    
    For Each Cell In .Rows(.Rows.Count).Cells
        
        Cell.Borders(3).weight = outBorderWidth
        Cell.Borders(3).ForeColor.RGB = outBorderColor
    
    Next Cell
    
    For Each Cell In .Columns(.Columns.Count).Cells
    
        Cell.Borders(4).weight = outBorderWidth
        Cell.Borders(4).ForeColor.RGB = outBorderColor
    
    Next Cell
End With
End If
                    
End Function

Public Sub CenterCell(ByRef Table, ByVal row, ByVal column)

    Table.Table.Cell(row, column).Shape.TextFrame2.HorizontalAnchor = msoAnchorCenter

End Sub

Public Sub FormatOneBorder(ByRef Table As Object, ByVal row As Integer, ByVal column As Integer, _
ByVal border As TableBorders, ByVal isVisible As Boolean, ByVal weight As Integer, ByVal color As Long)

With Table.Table.Cell(row, column).Borders(border)

    .Visible = isVisible
    .weight = weight
    .ForeColor.RGB = color
    
End With

End Sub

Public Sub LoadTextIntoCell(ByRef Table As Object, ByVal row As Integer, ByVal column As Integer, ByVal text As String, ByVal fontColor As Long, ByVal fontSize As Long, ByVal Append As Boolean, Optional ByVal font As String, Optional ByVal Bold As Boolean = False)

With Table.Table.Cell(row, column).Shape.TextFrame.TextRange
    If Append = False Then
    
            .text = text
            .font.color.RGB = fontColor

    Else
    
            .text = .text & " " & text
    
    End If

    .font.Size = fontSize
    .font.Name = font
    .font.Bold = Bold
End With
End Sub
Public Sub LoadRangeIntoCell(ByRef Table As Object, ByVal row As Integer, ByVal column As Integer, ByVal Cell As Range, Optional ByVal fontSize As Long, Optional ByVal font As String)

Cell.Copy
With Table.Table.Cell(row, column)

    .Shape.TextFrame2.TextRange.PasteSpecial 3 ' msoClipboardFormatRTF
    If fontSize <> 0 Then .Shape.TextFrame.TextRange.font.Size = fontSize
    If font <> "" Then .Shape.TextFrame.TextRange.font.Name = font

End With

End Sub
Public Sub ColorPartOfTheText(ByRef Table As Object, ByVal row As Integer, ByVal column As Integer, ByVal firstChar As Long, ByVal length As Long, ByVal color As Long)

Table.Table.Cell(row, column).Shape.TextFrame.TextRange.Characters(firstChar, length).font.color.RGB = color

End Sub

Public Sub SplitCell(ByRef Table As Object, ByVal row As Integer, ByVal Columm As Integer, ByVal NewRows As Integer, ByVal NewColumns As Integer)

Table.Table.Cell(row, Columm).Split NewRows, NewColumns

End Sub

Public Sub AddTextBox(ByVal slideID As Integer, ByVal textValue As String, ByVal fontFamily As String, ByVal fontSize As Integer, ByVal rotation As Double, _
ByVal color As Long, ByVal top As Integer, ByVal left As Integer, ByVal height As Integer, ByVal width As Integer, Optional ByVal orientation As powerPointAligement)

Dim textBox As Object
slideID = Builder(slideID)

Set textBox = currentPresentation.Slides(slideID).Shapes.AddTextBox(msoTextOrientationHorizontal, top, left, height, width)

With textBox
    .rotation = rotation
    .TextFrame2.TextRange.text = textValue
    .TextFrame2.TextRange.font.Name = fontFamily
    .TextFrame2.TextRange.font.Size = fontSize
    .TextFrame2.TextRange.font.Fill.ForeColor.RGB = color
    .ZOrder msoBringToFront
    If orientation <> 0 Then
    
        .TextFrame.TextRange.ParagraphFormat.Alignment = orientation
        
    End If
        
End With

End Sub

Sub AutoSizeTableRows(ByRef Table As Object, ByVal columnIndex As Integer)

Dim minH As Double

For i = 1 To Table.Table.Rows.Count

    With Table.Table.Cell(i, columnIndex).Shape.TextFrame
        
        minH = .TextRange.BoundHeight + .MarginTop + .MarginBottom
    End With
    Table.Table.Rows(i).height = minH
    
    
Next i

End Sub


Public Sub AdjustTableRows(ByRef Table, Heights() As Variant)

Dim finalArray() As Variant
Dim i As Integer
ReDim finalArray(1 To Table.Table.Rows.Count)
If Table.Table.Rows.Count > UBound(Heights) + 1 Then
    For i = 1 To UBound(Heights)
    
        finalArray(i) = Heights(i)
        
    Next i
    For j = i To Table.Table.Rows.Count
     
        finalArray(j) = Heights(i - 1)
        
    Next j
ElseIf Table.Table.Rows.Count < UBound(Heights) Then

    Debug.Print "Cound not allocate values to the array"

Else

    finalArray = Heights

End If

i = 1
For Each element In finalArray

    With Table.Table
    
        .Rows(i).height = element
    End With
    i = i + 1
Next element

End Sub


Public Function AddBackGroundToSlideMaster(ByVal slideID As Long, ByVal screenFill As Double, color As Long) As Object

slideID = Builder(slideID)
Dim top, left, width, height As Double
Dim slideWidth, slideHeight As Double
Dim thisShape As Variant

With currentPresentation.PageSetup

    slideHeight = .slideHeight
    slideWidth = .slideWidth

End With

width = slideWidth * screenFill
height = slideHeight * screenFill
top = (slideHeight - height) / 2
left = (slideWidth - width) / 2

Set thisShape = currentPresentation.Slides(slideID).Master.Shapes.AddShape(1, left, top, width, height)
thisShape.Fill.ForeColor.RGB = color
thisShape.Line.Visible = msoFalse
End Function

Public Sub AddLine(ByVal slideID As Integer, ByVal top1 As Long, ByVal left1 As Long, ByVal top2 As Long, ByVal left2 As Long, lineStyle As MsoLineDashStyle, ByVal color As Long, ByVal weight As Integer)

slideID = Builder(slideID)

With currentPresentation.Slides(slideID).Shapes.AddLine(top2, left2, top1, left1).Line

    .DashStyle = lineStyle
    .ForeColor.RGB = color
    .weight = weight
    
End With

End Sub
Public Sub AddRectangle(ByVal slideID As Integer, ByVal top As Long, ByVal left As Long, ByVal height As Long, ByVal width As Long, _
color As Long)
Dim thisShape As Object
slideID = Builder(slideID)
Set thisShape = currentPresentation.Slides(slideID).Shapes.AddShape(msoShapeRectangle, left, top, width, height)

With thisShape

    .Fill.ForeColor.RGB = color
    .Line.Visible = msoFalse

End With

End Sub

Private Function Builder(slideID) As Long
If currentPresentation Is Nothing Then

    Set currentPresentation = OpenOrCreatePowerPoint
    
End If

If currentPresentation.Slides.Count = 0 Then

    AddSlide
    
End If
    
If currentPresentation.Slides.Count < slideID Then

    slideID = currentPresentation.Slides.Count
    
End If
Builder = slideID

End Function
