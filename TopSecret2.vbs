Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the script's directory (your USB drive)
strScriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Define the Notepad file name
Dim notepadFileName
notepadFileName = "WifiProfilesWithPasswords.txt"

' Build the full path to the Notepad file
strFilePath = objFSO.BuildPath(strScriptDir, notepadFileName)

' Create a new text file for output
Set objFile = objFSO.CreateTextFile(strFilePath, True)

' Run the command to get all Wi-Fi profiles
Set objExec = objShell.Exec("cmd /c netsh wlan show profiles")
Do While Not objExec.StdOut.AtEndOfStream
    strLine = objExec.StdOut.ReadLine()
    If InStr(strLine, "All User Profile") > 0 Then
        ' Extract the Wi-Fi profile name
        wifiName = Trim(Split(strLine, ":")(1))
        ' Write the Wi-Fi profile name to the file
        objFile.WriteLine("Profile Name: " & wifiName)
        
        ' Run the command to show the Wi-Fi password for the current profile
        Set objExecProfile = objShell.Exec("cmd /c netsh wlan show profiles name=""" & wifiName & """ key=clear")
        Do While Not objExecProfile.StdOut.AtEndOfStream
            strProfileLine = objExecProfile.StdOut.ReadLine()
            If InStr(strProfileLine, "Key Content") > 0 Then
                ' Extract and write the Wi-Fi password to the file
                wifiPassword = Trim(Split(strProfileLine, ":")(1))
                objFile.WriteLine("Password: " & wifiPassword)
            End If
        Loop
        objFile.WriteLine("") ' Add a blank line between profiles
    End If
Loop

' Close the file
objFile.Close()

' Open Notepad with the Wi-Fi profiles and passwords
objShell.Run "notepad.exe " & strFilePath
