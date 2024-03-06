[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$uri = "$env:S1SiteURI"

$apiUri = "/web/api/v2.1/sites"

$body = @{
    apiToken = "$env:S1ReportingApiToken"
    sortBy = "name"
}

$uri += $apiUri

# Create an array to hold site data
$sitesData = @()

$response = Invoke-RestMethod -Uri $uri -Body $body

do {
    # Access sites data
    $sites = $response.data.sites

    # Loop through each site and extract required data
    foreach ($site in $sites) {
        $siteData = [PSCustomObject]@{
            Site = $site.name
            ActiveLicenses = $site.activeLicenses
            TotalLicenses = $site.totalLicenses
        }

        # Add site data to the array
        $sitesData += $siteData
    }

    $nextCursor = $response.pagination.nextCursor
    $body = @{
        apiToken = "$env:S1ReportingApiToken"
        cursor = $nextCursor
        sortBy = "name"
    }
    $response = Invoke-RestMethod -Uri $uri -Body $body
} while ($null -ne $nextCursor)

# Check if the directory exists, if not, create it
$directory = "C:\Temp"
if (!(Test-Path $directory)) {
    New-Item -Path $directory -ItemType Directory
}

# Export data to CSV
$sitesData | Export-Csv -Path "C:\Temp\sites_data.csv" -NoTypeInformation
