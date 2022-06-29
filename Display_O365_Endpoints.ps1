############################################# Load O365 JSON in a Powershell variable #######################################################
#
# You can either load the O365 JSON from a local file which download link is below for the O365 Worldwide endpoints:
# https://docs.microsoft.com/en-us/microsoft-365/enterprise/urls-and-ip-address-ranges
#
# All other O365 endpoints (US government, 21Vianet managed) including the worldwide endpoints links are on the below MS Site:
# https://docs.microsoft.com/en-us/microsoft-365/enterprise/microsoft-365-endpoints
#
# you can also directly download the latest O365 JSON file content and directly store it in a variable in Powershell
# to work with - that's  what we do in the below - please update the $Microsoft JSONURL according to your situationn,
# whether you need the O365 endpoints for the US Government (DoD, GCC), or using 21Vianet (sese link above for the other
# O465 endpoints JSON pointer.
#
##############################################################################################################################################


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

#region Script organizing O365 endpoints
# Here is a scriptlet that puts each URL and IP entry on a row (line), and for each URL and IP we show the corresponding port,
#whether it's required or not, and the notes for some URL and IPs (the non mandatory ones have a note).

# We'll use a PSCustomObject for this - for each IP or URL browsed, we create an item in our collection of PSCustomobjects with all the 
#other information (port, mandatory, notes,...) duplicated for each URL and IP:

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
#endregion

Write-Host "Number of IPs and URLs on expanded object: $($MyCollection.count)" -ForegroundColor Green
Write-Host "Type $MyCollection | ft IP, TCPPorts,UDPPorts, ServiceAreaDisplayName for example to see the object:"

$MyCollection | Select IP,URL, TCPPorts,UDPPorts, ServiceAreaDisplayName, required,notes | out-Gridview 
