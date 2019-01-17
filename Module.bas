Attribute VB_Name = "Module"
'Docs for VBA properties:
'   https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/freespace-property
'
' Modules to add in Tools/References:
'   Microsoft Scripting Runtime
Public Sub temp()
    
    fs = FreeFile()
    'Change .txt to preferred file type
    Dim path As String: path = "./log.txt"
    Dim pathName As String: pathName = Right(path, 7)
    Debug.Print pathName
    Dim world As String: world = "World"
    Dim name As Variant: name = InputBox("Kimi no nawa?")
    Dim fso As Object: Set fso = New FileSystemObject
    
    'Creating file in declared path,
    'If the file does exist then it will be overwritten.
    'To keep data of the existing file use 'Append' instead.
    Open path For Output As fs
    'When writing data, ending with semicolon means
    'continue writing in the same line.
        Print #fs, "Hello ";
        Print #fs, world; "!"
        Print #fs, "&"
        Print #fs, "Hi "; name & Chr(10) & Chr(13)
        Print #fs, "Lorem ipsum dolor sit amet,";
        Print #fs, "consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua."
        Print #fs, "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat."
        Print #fs, "Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur."
        Print #fs, "Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."
    Close #fs
    
    'Check if file existing in path.
    If fso.FileExists(path) Then
        MsgBox ("File has been created, " + name & Chr(13) & _
        Chr(10) & " Kindly check: " + pathName)
    Else
        MsgBox "File not created!"
    End If

End Sub

