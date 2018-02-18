Private Function Open_Browser(sWebsite As String, Optional bCloseBrowser As Boolean = False) As Boolean

 'bCloseBrowser set to FALSE means to OPEN BROWSER. TRUE = Close Browser

 'Make sure you've set a reference to the
 'Microsoft Internet Controls object library first
 '
 ' This code is from Wise Owl : http://www.wiseowl.co.uk/blog/s324/vba-ie.htm
 
 'create a variable to refer to an IE application, and
 'start up a new copy of IE (you could use GetObject to access an existing copy of you already had one open)
  
  
 'Check to see if the website is already open
  Dim ShellWins As SHDocVw.ShellWindows
  Dim ieApp As SHDocVw.InternetExplorer
  Dim WebBrowser As SHDocVw.WebBrowser
  
 'Instatiate object
  Set ShellWins = New SHDocVw.ShellWindows
  Set ieApp = New SHDocVw.InternetExplorer

'Set default value
  Open_Browser = False
 

On Error GoTo ProcErr
 
 If bCloseBrowser = False Then
 
    'make sure you can see this new copy of IE!
     With ieApp
     
        .Visible = True
        .Navigate sWebsite
           
     End With
     
    'wait for page to finish loading
     Do While ieApp.Busy And Not ieApp.ReadyState = READYSTATE_COMPLETE

        DoEvents

     Loop
     
 Else
 
    'make sure you can see this new copy of IE!
    'loop through browser to check for PreSalesDB
     For Each WebBrowser In ShellWins
    
        'Get address of webpages in browser
         sLocation = Left(WebBrowser.LocationURL, Len(sWebsite))
         
        'Check PreSales DB Site
         If sLocation = sWebsite Then
         
            WebBrowser.Quit
            Exit For
            
         End If
         
    Next
     
 
 End If
 
'Website
  Open_Browser = True
  

ProcExit:

 Set ShellWins = Nothing
 Set ieApp = Nothing
 Set WebBrowser = Nothing

    
Exit Function

ProcErr:
 
 'Set open website to false
  Open_Browser = False

  Select Case Err.Number
  
    Case 91  'Object not found Note: This occurs on the rsTrackChanges close statement
      'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Resume Next
      
      
    Case -2147467259 'Unspecified Error
    
      MsgBox "Error with opening Browser to Forecast Tool site!" & vbCrLf & vbCrLf & _
            "Send email to the itopursuitsites@atos.net mailbox for assistance", vbInformation + vbOKOnly, _
            "Function: Open_Browser Module: WinHTTP"
    
      Debug.Print "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
      Resume ProcExit
      
      
    Case -2147417848 'Object invoked has disconnecte from the client
    
      MsgBox "Error with opening Browser to Forecast Tool site!" & vbCrLf & vbCrLf & _
            "Send email to the itopursuitsites@atos.net mailbox for assistance", vbInformation + vbOKOnly, _
            "Function: Open_Browser Module: WinHTTP"
    
      Debug.Print "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
      Resume ProcExit
      

    Case -2147012889 'Server name address can not be resolved

      '    MsgBox "Forecast Tool PreSales DB site is not found!" & vbCrLf & vbCrLf & _
      '            "Check your connection to the internet of open the site at the address below" & vbCrLf & vbCrLf & _
      '             PUBLIC_URL_PRESALES, vbExclamation + vbOKOnly, "Function: Open_Browser Module: Mod_WinHTTP"
      
      Debug.Print "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
      Resume ProcExit


    Case Else
      MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
      Resume ProcExit
    
  End Select
    
Resume ProcExit

End Function