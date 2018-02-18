Private Function CheckURL_OpenInBrowser(sWebsite As String) As Boolean


'Check to see if the website is already open
 Dim ShellWins As SHDocVw.ShellWindows
 Dim Explorer As SHDocVw.InternetExplorer
 Dim WebBrowser As SHDocVw.WebBrowser
 
'Local variable
 Dim sLocation As String

'Set default value
 CheckURL_OpenInBrowser = False

On Error GoTo ProcErr

'Instatiate object
 Set ShellWins = New SHDocVw.ShellWindows


'loop through browser to check for PreSalesDB
 For Each WebBrowser In ShellWins

    'Get address of webpages in browser
     sLocation = Left(WebBrowser.LocationURL, Len(sWebsite))
     
    'Check PreSales DB Site
     If sLocation = sWebsite Then
     
    
        CheckURL_OpenInBrowser = True
        Exit For
        
    End If
    
 Next

  
ProcExit:

    Set ShellWins = Nothing
    Set Explorer = Nothing
    Set WebBrowser = Nothing

Exit Function

ProcErr:

  Select Case Err.Number
  
    Case 91  'Object not found Note: This occurs on the rsTrackChanges close statement
      'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Resume Next
      
    Case -2147012894 'Operation Timed Out

      MsgBox "Could Not Connect to the Forecast Tool PreSales DB!" & vbCrLf & vbCrLf & _
              "Check your connection to the internet of open the site at the address below" & vbCrLf & vbCrLf & _
              PUBLIC_URL_PRESALES, vbInformation + vbOKOnly
      
      Debug.Print "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
      Resume ProcExit
      
    Case -2147012889 'Server name address can not be resolved

      MsgBox "Check your connection to the internet of open the site at the address below", vbExclamation + vbOKOnly
      
      Debug.Print "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
      Resume ProcExit

    Case Else
      MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
      Resume ProcExit
    
  End Select
    
Resume ProcExit
    
    
End Function