Set WshShell = CreateObject("Wscript.Shell")

strLink1 = "https://github.com/choitaewon649/save/raw/main/1235.docx"

DesktopPath = WshShell.SpecialFolders("Desktop")
DownloadPath = WshShell.ExpandEnvironmentStrings("%userprofile%") & "\Downloads"


Set objFSO = CreateObject("Scripting.FileSystemObject")

   strSaveName1 = Mid(strLink1, InStrRev(strLink1,"/") + 1, Len(strLink1))
   strSaveTo1 = DownloadPath & "\" & strSaveName1
   strSaveTo2 = DesktopPath & "\" & strSaveName1

     ' Create an HTTP object
     Set objHTTP = CreateObject( "WinHttp.WinHttpRequest.5.1" )
 
     ' Download the specified URL
     objHTTP.Open "GET", strLink1, False
     ' Use HTTPREQUEST_SETCREDENTIALS_FOR_PROXY if user and password is for proxy, not for download the file.
     ' objHTTP.SetCredentials "User", "Password", HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
     objHTTP.Send
     
   'Set objFSO = CreateObject("Scripting.FileSystemObject")
   If objFSO.FileExists(strSaveTo1) Then
    objFSO.DeleteFile(strSaveTo1)
   End If

   If objFSO.FileExists(strSaveTo2) Then
    objFSO.DeleteFile(strSaveTo2)
   End If
 
      If objHTTP.Status = 200 Then
     Dim objStream
     Set objStream = CreateObject("ADODB.Stream")
     With objStream
      .Type = 1 'adTypeBinary
      .Open
      .Write objHTTP.ResponseBody
      .SaveToFile strSaveTo1
      .SaveToFile strSaveTo2
      .Close
     End With
     set objStream = Nothing
   End If
