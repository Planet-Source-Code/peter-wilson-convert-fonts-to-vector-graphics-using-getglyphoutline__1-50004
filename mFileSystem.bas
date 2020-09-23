Attribute VB_Name = "mFileSystem"
Option Explicit

' Used for writing temporary files
Declare Function GetTempFileName Lib "kernel32.dll" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Function GetWindowsTempFile() As String

    ' Generate a temporary file (path)\api????.TMP, where (path)
    ' is Windows's temporary file directory and ???? is a randomly assigned unique value.
    ' Then display the name of the created file on the screen.
    Dim temppath As String  ' receives name of temporary file path
    Dim tempfile As String  ' receives name of temporary file
    Dim slength As Long  ' receives length of string returned for the path
    Dim lastfour As Long  ' receives hex value of the randomly assigned ????
    
    ' Get Windows's temporary file path
    temppath = Space(255)  ' initialize the buffer to receive the path
    slength = GetTempPath(255, temppath)  ' read the path name
    temppath = Left(temppath, slength)  ' extract data from the variable
    
    ' Get a uniquely assigned random file
    tempfile = Space(255)  ' initialize buffer to receive the filename
    lastfour = GetTempFileName(temppath, "mdr", 0, tempfile)  ' get a unique temporary file name

    ' (Note that the file is also created for you in this case.)
    tempfile = Left(tempfile, InStr(tempfile, vbNullChar) - 1)  ' extract data from the variable
    
    GetWindowsTempFile = tempfile
    
End Function

