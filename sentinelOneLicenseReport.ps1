[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$uri = "$env:S1SiteURI"

$apiUri = "/web/api/v2.1/sites"

$body = @{
    apiToken = “$env:S1ReportingApiToken”
    sortBy = "name"
}

$uri += $apiUri

$response = Invoke-RestMethod -Uri $uri -Body $body

do {
    # Access sites data
    $sites = $response.data.sites

    # Loop through each site and extract required data
    foreach ($site in $sites) {
        $accountName = $site.name
        $activeLicenses = $site.activeLicenses
        $totalLicenses = $site.totalLicenses

        # Output required data
        Write-Output "Site: $accountName"
        Write-Output "Active Licenses: $activeLicenses"
        Write-Output "Total Licenses: $totalLicenses"
        Write-Output ""
    }

    $nextCursor = $response.pagination.nextCursor
    $body = @{
        apiToken   = “$env:S1ReportingApiToken”
        cursor = $nextCursor
        sortBy = "name"
    }
    $response = Invoke-RestMethod -Uri $uri -Body $body
} while (
    $null -ne $nextCursor
)
