Public Function FN_CheckUrl_Status(sWebsite As String) As Boolean

 ' This function checks if a website is running by sending an HTTP request.
 ' If the website is up, the function returns True, otherwise it returns False.
 ' Argument: sWebsite [string] in "www.domain.tld" format, without the
 ' "http://" prefix.
 '
 ' Written by Rob van der Woude
 ' http://www.robvanderwoude.com
 '
 ' Script site found : http://www.robvanderwoude.com/vbstech_internet_website.php
 
 ' XML DOM
 '
 ' Beginner guide: https://msdn.microsoft.com/en-us/library/aa468547.aspx

'WinHTTP Objects
 'Dim objHTTP As WinHttp.WinHttpRequest
 Dim objXMLHTTP As MSXML2.ServerXMLHTTP


'Local variable
 Dim iHTTP_Status As Integer
 Dim sHTTP_TextStatus As String
 
'Set default value
 FN_CheckUrl_Status = False
    
On Error GoTo ProcErr

'Instantiate objects
  'Set objHTTP = New WinHttp.WinHttpRequest
  Set objXMLHTTP = New MSXML2.ServerXMLHTTP
  

' With objHTTP
 With objXMLHTTP
 
    .Open "GET", sWebsite, False
    .SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MyApp 1.0; Windows NT 5.1)"
    .Send

 End With

'Set *** PUBLIC VARIABELS *** Status value and Status text
 'PUBLIC_WinHTTP_Status = objHTTP.Status
 'PUBLIC_WinHTTP_StatusText = objHTTP.StatusText
 
 PUBLIC_WinHTTP_Status = objXMLHTTP.Status
 PUBLIC_WinHTTP_StatusText = objXMLHTTP.StatusText
 
 
'Set Function value to TRUE
 If PUBLIC_WinHTTP_Status = 200 Then

    FN_CheckUrl_Status = True
    GoTo ProcExit
    
 End If


ProcExit:
    
 Set objHTTP = Nothing
 Set objXMLHTTP = Nothing
    
Exit Function

ProcErr:

  Select Case Err.Number
  
    Case 91  'Object not found Note: This occurs on the rsTrackChanges close statement
      'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Resume Next
      
    Case -2147012889 'Server name address can not be resolved
    
      PUBLIC_WinHTTP_Status = 111
      PUBLIC_WinHTTP_StatusText = "Not Connected to internet"
      
      Debug.Print "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
      Resume ProcExit
      
    Case -2147012890 'Site cant be found
    
      PUBLIC_WinHTTP_Status = 222
      PUBLIC_WinHTTP_StatusText = "Site Cant be found"
      
      Debug.Print "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
      Resume ProcExit
    
    Case -2147012739 'Site cant be found
      
      Debug.Print "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
      Resume ProcExit
    
    Case Else
    
      MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
      Debug.Print Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & Err.Source
      
      Resume ProcExit
    
  End Select
    
Resume ProcExit

End Function