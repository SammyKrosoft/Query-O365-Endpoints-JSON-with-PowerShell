# Display O365 Endpoints in a PowerShell table

Here's a script that displays the information from O365 enpoints into a nice table. It connects to the JSON file from Microsoft, and formats it as a table that you can paste in Excel, use in documentation, etc...


https://learn.microsoft.com/en-us/microsoft-365/enterprise/urls-and-ip-address-ranges?view=o365-worldwide


When you have PowerShell ISE (Integrated Scripting Environment), you can use the PowerShell GridView to display data, filter, sort, etc..., like a mini-Excel application. You can even select some (CTRL + click the valaues you want) or all the data in the table (CTRL + A), and copy these in an Excel or Word document.

In this repository, the *Display_O365_Endpoints.ps1* script does just that:
- Load the latest O365 list directly from the Microsoft website
- Parse all O365 entries and display these in a PowerShell GridView

![image](https://user-images.githubusercontent.com/33433229/176457473-f5fc4b73-bc6f-4597-93e7-11af727af495.png)

# Download Script

[Right Click "Save Link As" to download the script](https://raw.githubusercontent.com/SammyKrosoft/Query-O365-Endpoints-JSON-with-PowerShell/main/Display_O365_Endpoints.ps1)
