' ======================
' cygwin_update.vbs
' Desc: Download the last cygwin64 installer and execute it for update
' Vers: 0.4 (2015-05-17T20:00:02+01:00)
' Author: Tony P
' ======================
' CHANGELOG:
'   - 0.4
'       - ADDED
'           - check version number before downloading setup  

'   - 0.3
'       - ADDED
'          - Determine online version number (https://sourceware.org/cgi-bin/cvsweb.cgi/setup/?cvsroot=cygwin-apps  need to extract ChangeLog file Rev number)
' TODO:
'   - ADD
'       - Skip when check was done under a given delay
'       - Run setup in silent mode
'       - read output
'   - 
'
Option Explicit
dim setupGitPage, lastVerNum, localVerNum
dim iOSBits, sFilename, sDownDir, sURL, sDownLoc, sVersionFile
dim oXHttp, oAdoStr, objFSO, objFile, oWShell, oExec, oRun


forceCScriptExecution ' stop & run this script in cscript if needed

WScript.Echo "===============================" _
&vbNewline  &"Cygwin update script" _
&vbNewline  &"===============================" _
&vbNewline

' Set variables depending arch
iOSBits=OSBits()
select case iOSBits
    case 32
        sDownDir = "C:\cygwin\"
        sFileName = "setup.exe"
    case 64
        sFileName = "setup-x86_64.exe"
        sDownDir = "C:\cygwin64\"
    case else    
        WScript.Echo "Error: Attemp to determine arch failed" _
                      &vbNewline &"The script will stop in few seconds"                
        WScript.Sleep 8000
        WScript.Quit
end select
WScript.Echo iOSBits &" bits system detected."

sURL = "http://cygwin.com/"&sFileName
sDownLoc = sDownDir&sFileName

' Check online version number
'Set setupGitPage = WScript.GetObject("https://sourceware.org/viewvc/cygwin-apps/setup/")
Set setupGitPage = WScript.GetObject("https://cygwin.com/git/gitweb.cgi?p=cygwin-setup.git;a=tags")
While setupGitPage.readystate <> "complete" 
    WScript.Echo "." 
    WScript.Sleep 500 
Wend 
WScript.Echo TypeName(setupGitPage)

On Error Resume Next
lastVerNum = trim(setupGitPage.getElementsByTagName("table")(0).firstChild.getElementsByTagName("td")(1).InnerText)
lastVerNum = Replace(lastVerNum, "release_", "")
lastVerNum = Replace(lastVerNum, ".", ",")
lastVerNum = CDbl(lastVerNum)
WScript.Echo lastVerNum
' If all is ok
If Err.Number = 0 Then
	WScript.Echo "Online version number: " &lastVerNum
Else
	WScript.Echo vbTab&"** Error Num: " &Err.Number &" Src: " &Err.Source &" Desc: " &Err.Description
	Err.Clear
    WScript.Echo "Cannot determine the online version number, script will stop"
    WScript.Sleep 5000
    WScript.Quit
End if
On Error Goto 0

' Check local version number
sVersionFile = sDownDir&"localVersion.txt"
' we create FSo object for reading version number previously saved
On Error Resume Next
set objFSO  = createObject("Scripting.FileSystemObject")
set objFile = objFSO.OpenTextFile(sVersionFile)
' Check locale version number
localVerNum = CDbl(objFile.ReadAll)
objFile.Close

' If all is ok
If Err.Number = 0 Then
	WScript.Echo "Local version number: " &localVerNum
Else
	WScript.Echo vbTab&"** Error Num: " &Err.Number &" Src: " &Err.Source &" Desc: " &Err.Description
	Err.Clear
End if
On Error Goto 0


if lastVerNum > localVerNum then
    WScript.Echo "Online version more recent"

    if downloadFile(sURL, sDownLoc) then
        ' we create a file contening version number
        Set objFile = objFSO.CreateTextFile(sVersionFile,True)
        objFile.Write lastVerNum & vbCrLf
        objFile.Close
     
        ' set Working Directory (usefull for saving here, logfiles created by cygwin setup)
        set oWShell = createObject("WScript.Shell")
        oWShell.CurrentDirectory = sDownDir

        ' Run donwloaded file (cygwin setup)
        WScript.Echo vbNewLine&vbTab&" installer will be execute"
        WScript.Sleep 3000
        'Set oExec = oWShell.Exec(sDownLoc &" --quiet-mode --no-desktop")
        Set oExec = oWShell.Exec(sDownLoc &" -q -g -d")
        Do While oExec.Status = 0
            WScript.Sleep 500
        Loop
        WScript.Echo "Status " & oExec.Status
        set oExec = nothing
        set oWShell = nothing

    else
        WScript.Echo "Error during file download attempt"
    end if
elseif lastVerNum = localVerNum then
    WScript.Echo "Locale version up to date"
    WScript.Sleep 8000
else
    WScript.Echo "Issue during version numbers comparison, please check"
    WScript.Sleep 15000
    WScript.Quit
end if

'Pause("Press any key to continue")


Function downloadFile(sURL, sDownLoc)

    set oXHttp  = createObject("Microsoft.XMLHTTP")
    set oAdoStr = createObject("Adodb.Stream")

    ' Request URL
    oXHttp.Open "GET", sURL, false
    oXHttp.Send

    WScript.Echo vbNewLine&"Cygwin download attempt:"
    WScript.Echo vbTab&"- URL: "&sURL
    WScript.Echo vbTab&"- Saving path: "&sDownLoc


    ' if URL found
    if oXHttp.Status = 200 Then
        WScript.Echo vbNewLine&vbTab&"- URL finded"

        On Error Resume Next
        ' Save response to file
        with oAdoStr
            .type = 1 '//binary
            .open
            .write oXHttp.responseBody
            .savetofile sDownLoc , 2 '//overwrite
            .close
        end With
        
        ' If response saved
        If Err.Number = 0 Then
            WScript.Echo vbTab&"- File saved"
            downloadFile = 1
        Elseif Err.Number = 3004 then
            WScript.Echo vbTab&"** Error during file saving"
            downloadFile = 0
        Else
            WScript.Echo vbTab&"** Error Num: " &Err.Number &" Src: " &Err.Source &" Desc: " &Err.Description
            Err.Clear
            downloadFile = 0
        End if
        On Error Goto 0
        
    else
        WScript.Echo vbTab&"** Issue with url :" +sURL
        WScript.Echo vbTab&"** oXHttp.Status :" +oXHttp.Status
    end if

    Set oAdoStr = nothing
    Set oXHttp = nothing
End Function

' Force to execute in cscript
Sub forceCScriptExecution
    Dim oWShell
    Dim sArg, sStr, sMsg

    If Not LCase( Right( WScript.FullName, 12 ) ) = "\cscript.exe" Then
        set oWShell = createObject( "WScript.Shell" )
        For Each sArg In WScript.Arguments
            If InStr( sArg, " " ) Then sArg = """" & sArg & """"
            sStr = sStr & " " & sArg
        Next
        sMsg = "Detetcted Interpretor: " &WScript.FullName &vbNewLine &vbNewLine &"Script will be execute with cscript "
        if not isEmpty(sStr) then
            sMsg = sMsg &vbrc &vbNewLine &"Arguments: " &sStr
        end if
        oWShell.Popup sMsg, 3, "Script prévu pour être interpréter par cscript", 0
        oWShell.Run "cscript """ & WScript.ScriptFullName & """ " & sStr , 1, False
        WScript.Quit
    End If
    set oWShell = nothing
End Sub

' Pause fonctionnality : Display a message and wait for a pressed key
Sub Pause(strPause)
    dim z
    WScript.Echo (strPause)
    z = WScript.StdIn.Read(1)
End Sub

' Check architecture of the system
Function OSBits()
    OSBits = GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth
End Function