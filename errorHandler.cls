VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "errorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum errorTypes

    Warning
    StandardError
    CriticalError
    
End Enum


Private FSO As Object
Private errorDirectory As String
Private errorCollection As Collection
Private LogInFile As Range

Private Sub Class_Initialize()

    errorDirectory = ActiveWorkbook.Path
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set errorCollection = New Collection
    
End Sub
Private Sub Class_Terminate()

If errorCollection.Count > 0 Then

    GenerateTextLog
    
End If

Set FSO = Nothing
Set errorCollection = Nothing
Set LogInFile = Nothing

End Sub

Public Function SetLog(logRNG As Range)


If logRNG.Count > 1 Then

    Set LogInFile = logRNG.Cells(1, 1)

Else

    Set LogInFile = logRNG

End If


End Function

Public Function AddError(ByVal errorString As String, errorType As errorTypes)

Dim errorStart As String
Dim errorEnd As String
Dim exitFlag As Boolean
Dim finalString As String

Select Case errorType

    Case Warning
    
        errorStart = "Warning!: "
        errorEnd = "Application will continue, however contact the administrator."
        exitFlag = False
        
    Case StandardError
    
        errorStart = "Error!: "
        errorEnd = "Application will continue, however this error may result in data loss or incorrect calculations."
        exitFlag = False
        
    Case CriticalError
    
        errorStart = "Critical Error!: "
        errorEnd = "Application execution interrupted."
        exitFlag = True

End Select

finalString = "[" & Now & "] " & errorStart & errorString & " " & errorEnd

errorCollection.Add finalString
AddToLog finalString, errorType
If exitFlag Then


    GenerateTextLog
    MsgBox "The Macro encountered a critical error code execution interupted", vbCritical, "Critical Error"
    End
    
End If
    

End Function
Private Function AddToLog(ByVal errorString As String, errorType As errorTypes) As Boolean

Dim currentLogRange As Range
Dim logRow As Long
Dim logCol As Long

If Not LogInFile Is Nothing Then

    logCol = LogInFile.Column
    
    Set currentLogRange = LogInFile.CurrentRegion
    
    With currentLogRange
    
        logRow = .Rows(.Rows.Count).Row + 1
        
    End With
    Set currentLogRange = LogInFile.Parent.Cells(logRow, logCol)
    With currentLogRange
        .Value = errorString
        Select Case errorType
        
            Case Warning
            
                .Font.Color = RGB(240, 171, 0)
                
            Case StandardError
            
                .Font.Color = RGB(255, 0, 0)
                
            Case CriticalError
            
                .Font.Color = RGB(255, 0, 0)
                .Font.Bold = True
                
        End Select
    End With
End If
End Function

Private Function GenerateTextLog()

Dim file As Object
Dim errorFile As String

errorFile = errorDirectory & _
            "\ErrorLog " & _
            Format(Now, "YYYY-MM-DD HHMMSS") & _
            ".txt"

Set file = FSO.CreateTextFile(errorFile, True)

For Each element In errorCollection

    file.WriteLine element
    
Next element

file.Close


End Function


