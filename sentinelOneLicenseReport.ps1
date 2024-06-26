[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$uri = $env:S1SiteURI

$apiUri = "/web/api/v2.1/sites"

$apiToken = $env:S1ReportingApiToken

$body = @{
    apiToken = $apiToken
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
        apiToken = $apiToken
        cursor = $nextCursor
        sortBy = "name"
    }
    $response = Invoke-RestMethod -Uri $uri -Body $body
} while ($null -ne $nextCursor)

# Export data to CSV
$csvFilePath = ".\sites_data.csv"

#export to path
$sitesData | Export-Csv -Path $csvFilePath -NoTypeInformation

# Install Mailozaurr
Get-PackageProvider -Name NuGet -Force
Install-Module -Name Mailozaurr -AllowClobber -Force
Import-Module -Name Mailozaurr -Force

# Send email with CSV attachment
$smtpServer = $env:SMTPServer
$smtpPort = $env:SMTPPort
$from = $env:EmailSendFromAddress
$to = $env:S1ReportingDestinationEmail
$SMTPUsername = $env:SMTPEmailUsername
$SMTPPassword = $env:SMTPEmailPassword
[securestring]$secStringPassword = ConvertTo-SecureString $SMTPPassword -AsPlainText -Force
[pscredential]$EmailCredential = New-Object System.Management.Automation.PSCredential ($SMTPUsername, $secStringPassword)
$subject = "Sentinel One License Report"
#todo: Figure out how to get newlines to work in body of email. The below doesn't work. Still sends the body just fine, just without newlines. 
$body = @"
Please find attached the Sentinel One License Report CSV file.
If you have questions do not reply to this message, please send a message to the NOC in NOC-Toolkit or email $env:NOCEmail.
"@
$attachment = $csvFilePath

Send-EmailMessage `
    -SmtpServer $smtpServer `
    -Port $smtpPort `
    -From $from `
    -To $to `
    -Credential $EmailCredential `
    -Subject $subject `
    -Body $body `
    -Attachments $attachment
    
