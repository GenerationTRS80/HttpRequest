Private Sub WinHTTP_Status_Msg(iWinHTTP_Status As Integer)

 Select Case iWinHTTP_Status
 
  Case 111 'WiFi is off
  
          MsgBox "You are not connected to the internet. Check you internet connection" _
                  , vbCritical + vbOKOnly, "Function WinHTTP_Status_Msg Module: WinHTTP"
                  
  Case 222 'HTTP_STATUS_FORBIDDEN
  
        MsgBox "Site cant be found. Wrong URL" _
                  , vbExclamation + vbOKOnly, "Function Main Module: ForecastToolsReports"
                  
  Case 400 'HTTP_STATUS_BAD_REQUEST
  
        MsgBox "Forecast tool site has timed out. You need to reopen it" _
                  , vbExclamation + vbOKOnly, "Function Main Module: ForecastToolsReports"
                  
  Case 403 'HTTP_STATUS_FORBIDDEN
  
          MsgBox "You don't have permission to use SharePoint" _
                  , vbExclamation + vbOKOnly, "Function Main Module: ForecastToolsReports"
                  
  Case 404 'HTTP_STATUS_NOT_FOUND

        MsgBox "Site Not Found" _
                  , vbCritical + vbOKOnly, "Function Main Module: ForecastToolsReports"
  
  Case Else
  
          MsgBox "Error with connecting to Forecast Tool site!" & vbCrLf & vbCrLf & _
                  "Status " & PUBLIC_WinHTTP_StatusText & " " & PUBLIC_WinHTTP_Status & vbCrLf & vbCrLf & _
                "Send email to the itopursuitsites@atos.net mailbox for assistance", vbInformation + vbOKOnly, _
                "Function: WinHTTP_Status_Msg Module: WinHTTP"
 
 End Select


End Sub