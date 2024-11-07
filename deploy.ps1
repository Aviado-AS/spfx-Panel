

gulp clean
gulp bundle --ship
gulp package-solution --ship

function Deploy-AppToSites {
    param (
        [string[]]$sites,
        [string]$appPath,
        [string]$applicationId
    )

    foreach ($url in $sites) {
        Write-Host "Deploying to $url" -ForegroundColor "Green"
        Connect-PnPOnline -Url $url -Interactive -ApplicationId $applicationId
        Add-PnPSiteCollectionAppCatalog -Site $url -ErrorAction SilentlyContinue
        $app = Add-PnPApp -Path $appPath -Scope Site -Publish -Overwrite
        Install-PnPApp -Identity $app.Id -Scope Site -ErrorAction SilentlyContinue
    }
}




$appPath = "C:\GitHub\brg\SPFX\copyPage\sharepoint\solution\copypage.sppkg"
$applicationId = "376ea684-f179-466a-ad96-dd510b7a2e79"
Connect-PnPOnline -Url https://borregaard-admin.sharepoint.com  -Interactive -ApplicationId $applicationId

$globalHub = Get-PnPHubSite -Identity "https://borregaard.sharepoint.com" 

$globalHubSites = Get-PnPHubSiteChild -Identity $globalHub
$globalHubSites = $globalHubSites + "https://borregaard.sharepoint.com/"

Deploy-AppToSites -sites $globalHubSites -appPath $appPath -applicationId $applicationId

$noHub = Get-PnPHubSite -Identity "https://borregaard.sharepoint.com/sites/Portal-NO" 

$noHubSites = Get-PnPHubSiteChild -Identity $noHub
$noHubSites = $noHubSites + "https://borregaard.sharepoint.com/sites/Portal-NO" 

Deploy-AppToSites -sites $noHubSites -appPath $appPath -applicationId $applicationId


Copy-Item -Path $appPath -Destination "C:\GitHub\brg\SPFx packages" -Force
