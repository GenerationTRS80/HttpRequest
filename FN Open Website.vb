
Public Function FN_Open_Website(xlWrkSht_Button As Excel.Worksheet, sURL_website As String) As Boolean

 '--------------------------------------------------------------------------------------------
 '
 '  NOTE: This program
 '
 '
 
 'Excel Object
  Dim xlWrkBk_Forecast As Excel.Workbook
  Dim rngLoggedIn_Hours As Range
  
 'Local Variable
  Dim iLoggedIn_Hours As Integer
  Dim bUpdated_Website As Boolean
 
 'Set default value for function
  FN_Open_Website = False
  bUpdated_Website = False 'If the website has been opened by the Open_Website function set to true
    
    
On Error GoTo ProcErr

 'Instantiate Object
  Set xlWrkBk_Forecast = xlWrkSht_Button.Parent
  Set rngLoggedIn_Hours = xlWrkBk_Forecast.Names("LoggedIn_Hours").RefersToRange
 
  iLoggedIn_Hours = rngLoggedIn_Hours.Value
   
 'Remove Protection
  FN_Public_UnProtect_Workbook
 
'-----------------------------------------------------------------------------------------------
' Open PreSale DB website
'
' NOTE: 1) need to check if busy or DAS ID Required
'       2) Check if already open
'
 

'   Check to make sure that Forecast tool is already open
'   NOTE: Need  to check if open in browser first because this subroutine is fast and CheckUrl_Exist is slow
'         CheckUrl_Exist sub is slow if it is before CheckURL_OpenInBrowser
 If CheckURL_OpenInBrowser(sURL_website) = False Then
 
    'NOTE: **** There is problem with "An error occurred in the secure channel support msxml3.dll" *****

    'Check if URL Exists
     If FN_CheckUrl_Status(sURL_website) = False Then

        'Show Status message
          WinHTTP_Status_Msg PUBLIC_WinHTTP_Status

         'Exit sub
          GoTo ProcExit

     End If
 
    '**** OPEN Website ***
     If Open_Browser(sURL_website) = False Then
    
          MsgBox "Forecast Tool can NOT be Opened!" & vbCrLf & vbCrLf & _
                  "Send email to the itopursuitsites@atos.net mailbox for assistance", vbInformation + vbOKOnly, "Function Main Module: ForecastToolsReports"
                  
         'Exit sub
          GoTo ProcExit
          
     Else
     
        'If the website has been opened by the Open_Website function set to true
         bUpdated_Website = True
        
     End If
      
      
 Else
 
    'Check if logged in more that an Hour if more than an hour then reopen site ie reloggin to the site
     If iLoggedIn_Hours >= PUBLIC_LOGGEDIN_HOURS Then
    
        ' ***Close webpage ***
         If Open_Browser(sURL_website, True) = False Then
    
             MsgBox "Forecast Tool had and Error in closing!" & vbCrLf & vbCrLf & _
                  "Send email to the itopursuitsites@atos.net mailbox for assistance", vbInformation + vbOKOnly, "Function Main Module: ForecastToolsReports"
                  
            'Exit sub
             GoTo ProcExit
        
         End If
    

        '**** Re OPEN Website ***
         If Open_Browser(sURL_website) = False Then
    
            MsgBox "Forecast Tool can NOT be Opened!" & vbCrLf & _
                  "Send email to the itopursuitsites@atos.net mailbox for assistance", vbCritical + vbOKOnly, "Function Main Module: ForecastToolsReports"
                  
           'Exit sub
            GoTo ProcExit
        
         Else
     
           'If the website has been opened by the Open_Website function set to true
            bUpdated_Website = True
        
         End If
                
      Else

        'If the Forecast tool is ** Already Open ** AND it has been less than an hour since the last time logged in
        'THEN Check to make sure that the SharePoint site and or Internet is NOT disconnected
         If FN_CheckUrl_Status(sURL_website) = False Then
        
               'Show Status message
                WinHTTP_Status_Msg PUBLIC_WinHTTP_Status
                
               ' ***Close webpage ***
                If Open_Browser(sURL_website, True) = False Then
            
                    MsgBox "Forecast Tool had and Error in closing!" & vbCrLf & vbCrLf & _
                          "Send email to the itopursuitsites@atos.net mailbox for assistance", vbInformation + vbOKOnly, "Function Main Module: ForecastToolsReports"
                          
                    'Exit sub
                     GoTo ProcExit
                
                     End If
                       
               'Exit sub
                GoTo ProcExit
            
          Else
         
                MsgBox "Already Open in Browser!", vbInformation + vbOKOnly
           
         End If
         

      End If
      
 
 End If
 
 
'Set to FN_Open_Website true
 FN_Open_Website = True
 
'Set time of web page update if the Web site has been opened
 If bUpdated_Website Then
 
     With xlWrkBk_Forecast.Worksheets("Main")
    
       '.Activate
       .Range("D3").Select
       .Range("D3").Value = Now()
    
     End With
     
 End If
 
 
ProcExit:

'Protect Workbook
 FN_Public_Protect_Workbook


Exit Function

ProcErr:

 FN_Open_Website = False

  Select Case Err.Number
  
    Case 91  'Object not found Note: This occurs on the rsTrackChanges close statement
      'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Resume Next
      
    Case 1004 'Protection is set for the cell
      MsgBox "Cell is protected can not write value", vbInformation
      Resume Next

    Case Else
      MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
      Resume ProcExit
    
  End Select
    
Resume ProcExit
    
End Function