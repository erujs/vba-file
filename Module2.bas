Attribute VB_Name = "Module2"

Option Explicit

' Go to Tools -> References... and check "Microsoft Scripting Runtime" to be able to use
' the FileSystemObject which has many useful features for handling files and folders
Public Sub SaveTextToFile()

    Dim filePath As String
    filePath = ".\test.sas"

    ' The advantage of correctly typing fso as FileSystemObject is to make autocompletion
    ' (Intellisense) work, which helps you avoid typos and lets you discover other useful
    ' methods of the FileSystemObject
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim fileStream As TextStream

    ' Here the actual file is created and opened for write access
    Set fileStream = fso.CreateTextFile(filePath)

    ' Write something to the file
    fileStream.WriteLine "data copper;"
    fileStream.WriteLine "input id warp temp pct;"
    fileStream.WriteLine "datalines;"
    fileStream.WriteLine "1    17   50  40"
    fileStream.WriteLine "2    20   50  40"
    fileStream.WriteLine "3    16   50  60"
    fileStream.WriteLine "4    21   50  60"
    fileStream.WriteLine ";"
    fileStream.WriteLine "proc anova data=copper;"
    fileStream.WriteLine "  class temp pct;"
    fileStream.WriteLine "model warp= temp | pct;"
    fileStream.WriteLine "run;"

    ' Close it, so it is not locked anymore
    fileStream.Close

    ' Here is another great method of the FileSystemObject that checks if a file exists
    If fso.FileExists(filePath) Then
        MsgBox "Completed!"
    End If

    ' Explicitly setting objects to Nothing should not be necessary in most cases, but if
    ' you're writing macros for Microsoft Access, you may want to uncomment the following
    ' two lines (see https://stackoverflow.com/a/517202/2822719 for details):
    'Set fileStream = Nothing
    'Set fso = Nothing

End Sub
