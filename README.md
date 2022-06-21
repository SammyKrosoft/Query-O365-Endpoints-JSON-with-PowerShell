```powershell

############################################# Load O365 JSON in a Powershell variable #######################################################

# You can either load the O365 JSON from a local file (if you choose to download it from MS), or 
# you can also directly download the latest O365 JSON file and directly store it in a variable in Powershell
# to work with:
#region Load JSON file from local
$FileLocation = "c:\temp\O365ips.json"
$O365Endpoints = Get-Content $FileLocation | ConvertFrom-Json
#endregion

#region Load JSON directly from MS Website
# You can load directly the Microsoft O365 endpoints within Powershell (provided outgoing port 443 is opened on the station where you run the code)
# Here is to load from MS Website, the Try {} catch {} are for error control in case you ahve the wrong download URL:
$MicrosoftJSONURL = "https://endpoints.office.com/endpoints/worldwide?clientrequestid=b10c5ed1-bad1-445f-b386-b919946339a7"
$web_client = new-object system.net.webclient

Try{
    $O365Endpoints=$web_client.DownloadString($MicrosoftJSONURL) | ConvertFrom-Json

    Write-Host "URLs and IPs Loaded !" -ForegroundColor Green
    Write-Host "We found $($O365Endpoints.count) endpoints !"
    Write-host "to search the endpoints, use the `$O365Endpoints variable"
    }

Catch {
Write-Host "An error occurred, probably your URL is not the good one..." -ForegroundColor Green
Write-Host "URL used : $MicrosoftJSONURL" -ForegroundColor Green
Write-Host "Error : " -ForegroundColor Red -BackgroundColor Yellow
Write-Host "$($Error[0].Exception)" -ForegroundColor Red
$O365Endpoints = $null
}

#endregion


############################################################## Filters samples #############################################################

# Filter only Exchange online entries, dumping which are required, which are not and display as a table:
$O365Endpoints | ? {$_.ServiceAreaDisplayName -eq "Exchange Online"} | ft Id, Required, URLs, IPs, TCPPorts, UDPPorts
# NOTE: if your PowerShell windows buffer is not big enough, Powershell will only print the first 3 or 4 columns

# Filter Exchange Online entries only, and only these that are required:
$O365Endpoints | ? {$_.ServiceAreaDisplayName -eq "Exchange Online" -and $_.required -eq $True} | ft Id, Required, URLs, IPs, TCPPorts, UDPPorts
# NOTE: if your PowerShell windows buffer is not big enough, Powershell will only print the first 3 or 4 columns
 
# Display Exchange Online entries only, and only those that are required, but in List form 
# - that way you get all fields compared to usingthe | ft
$O365Endpoints | ? {$_.ServiceAreaDisplayName -eq "Exchange Online"-and $_.required -eq $True} | fl Id, Required, URLs, IPs, TCPPorts, UDPPorts

# To display all URLs all at once (without the ports):
$O365Endpoints | ? {$_.ServiceAreaDisplayName -eq "Exchange Online"-and $_.required -eq $True} | Select -ExpandProperty URLs -ErrorAction SilentlyContinue

# To Display all IPs all at once (without the ports):
$O365Endpoints | ? {$_.ServiceAreaDisplayName -eq "Exchange Online"-and $_.required -eq $True} | Select -ExpandProperty IPs -ErrorAction SilentlyContinue

#NOTE: you can store IPs and URLs List in a variable
# First store the URLs list:
$O365URLsList_ExO_Required = $O365Endpoints | ? {$_.ServiceAreaDisplayName -eq "Exchange Online"-and $_.required -eq $True} | Select -ExpandProperty URLs -ErrorAction SilentlyContinue
# Then store the IPs list:
$O365IPsList_ExO_Required = $O365Endpoints | ? {$_.ServiceAreaDisplayName -eq "Exchange Online"-and $_.required -eq $True} | Select -ExpandProperty IPs -ErrorAction SilentlyContinue

# You can either use each separately, or concatenate the 2 lists in one common list:
$O365URLs_IPs_ExO_required = $O365URLsList_ExO_Required + $O365IPsList_ExO_Required

# Then you can search a particular IP or URL on the list using -Like "*Outlook*" for example
$O365URLs_IPs_ExO_required -Like "*outlook*"

<#
Output will look like the below

outlook.office.com
outlook.office365.com
*.outlook.com
*.protection.outlook.com
*.mail.protection.outlook.com

#>



#region Below is an example putting all URLs and IPs on a single list to be able to check
# We will put just the URLs and IPs on a plain list of IPs and URLs in a PowerShell variable

     # First put the IPs in $IPList variable:
     $IPList = $O365Endpoints | select -ExpandProperty IPs -ErrorAction SilentlyContinue
     # Second, put the URLs in $URList variable:
     $URList = $O365Endpoints | select -ExpandProperty URLs -ErrorAction SilentlyContinue
     # Third, concatenate (join) the two above lists in one common list
     $O365Everything = $IPList + $URList

     # Then you can Search for an URL or an IP !
     # For example you're looking if the list has Linkedin in it:
     $O365Everything -like "*outlook*"

#endregion
 
 


# Here is a scriptlet that puts each URL and IP entry on a column, and the corresponding port on another , and whether it's required or not

# We'll use a PSCustomObject for this - for each IP or URL browsed, we create an item in our collection of PSCustomobjects with all the 
#other information duplicated for each URL and IP:

$MyCollection = @()
Foreach ($endpoint in $O365Endpoints) {
   If ($endpoint.ips){
        Foreach ($IP in $endpoint.ips){
            $MyObject = [PSCustomObject]@{
                ID = $endpoint.id
                ServiceArea = $endpoint.serviceArea
                ServiceAreaDisplayName = $endpoint.serviceAreaDisplayName
                IP = $IP
                URL = ""
                TCPPorts = $endpoint.tcpPorts
                UDPPorts = $endpoint.udpPorts
                Required = $endpoint.required
                Expressroute = $endpoint.expressRoute
                Notes = $endpoint.notes
            }
        $MyCollection += $MyObject
          }
       }
      If ($endpoint.urls.count -gt 0) {
        Foreach ($URL in $endpoint.urls){
            $MyObject = [PSCustomObject]@{
                ID = $endpoint.id
                ServiceArea = $endpoint.serviceArea
                ServiceAreaDisplayName = $endpoint.serviceAreaDisplayName
                IP = ""
                URL = $URL
                TCPPorts = $endpoint.tcpPorts
                URLPorts = $endpoint.udpPorts
                Required = $endpoint.required
                Expressroute = $endpoint.expressRoute
                Notes = $endpoint.notes
            }
        $MyCollection += $MyObject
          }
      }
}

Write-Host "Number of IPs and URLs on expanded object: $($MyCollection.count)" -ForegroundColor Green
Write-Host "Type $MyCollection | ft IP, TCPPorts,UDPPorts, ServiceAreaDisplayName for example to see the object:"
$MyCollection | ft IP,URL, TCPPorts,UDPPorts, ServiceAreaDisplayName

Write-Host "Type $MyCollection | Select IP, TCPPorts,UDPPorts, ServiceAreaDisplayName | out-Gridview to open on a table and copy/paste in Excel for example"

$MyCollection | Select IP,URL, TCPPorts,UDPPorts, ServiceAreaDisplayName, required,notes | out-Gridview 

```
