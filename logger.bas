Attribute VB_Name = "logger"
Option Compare Database
Option Explicit

Const folderName As String = "\Log"

Function logWrite(str As String)
Dim folder As String
folder = CurrentProject.path & folderName

Dim history As String
history = logRead ' logread is a function below

Dim location As String
location = folder & "\logger.log"
Open location For Output As #1
    Print #1, history
    Print #1, Format(Now, "h:mm:ss ampm") & " " & str
Close #1
End Function

Function logRead() As String

Dim folder As String
folder = CurrentProject.path & folderName
Dim location As String
Dim textline As String
location = folder & "\logger.log"
If DisplayFunction.FileFolderExists(folder) = False Then
    'MsgBox "Log Folder is missing, a new folder will be created!"
    MkDir folder
End If

If DisplayFunction.FileExists(location) = True Then
    Dim text As String
    
    Open location For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            
            If Len(text) > 0 Then
                text = text & vbNewLine & textline
            Else
                text = text & textline
            End If
        Loop
    Close #1
    logRead = text
Else
    logRead = Format(Now, "h:mm:ss ampm") & " Log Created."
End If
End Function
