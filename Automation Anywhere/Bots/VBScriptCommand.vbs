Dim Fso, FileObj, FilePath
Set Fso = CreateObject("Scripting.FileSystemObject") 'Creates "FileSystemObject" Object.
FilePath = WScript.Arguments.Item(0)
    'MsgBox  FilePath' 'Uncomment this line to see the filename'
    Set FileObj = Fso.GetFile(FilePath)
    WScript.Echo FileObj.DateCreated