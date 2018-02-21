# HttpRequest
This is an application module within the SharePoint-ReportingApp. The module will open an website in IE. If the website is already open it will then reload the website.

VBScript using WinHttpRequest and XMLHttpRequest to accessing the website.
Microsoft Internest Controls (VBA) to manipulate IE 


The scripts will do the following:
1 - Check the browser to see if the webpage is already open
2 - Check the status of the webpage if it is open
2a. - If the webpage has been open for more than 2 hours then reopen it
3 - Open the webpage if the page is not found in the browser

Note: Make sure in Excel you have these libraries set in references 
  tools -> references -> Microsoft WinHTTP Services, version 5.1,
                         Microsoft Internet Controls,
                         Microsoft Scripting Runtime,
                         Microsoft XML v6.0

