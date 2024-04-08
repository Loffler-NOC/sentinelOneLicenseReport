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
$csvFilePath = "C:\Temp\sites_data.csv"
# Check if C:\Temp directory exists
if (-not (Test-Path -Path "C:\Temp" -PathType Container)) {
    # Create C:\Temp directory if it doesn't exist
    New-Item -Path "C:\Temp" -ItemType Directory
    Write-Host "C:\Temp directory created successfully."
} else {
    Write-Host "C:\Temp directory already exists."
}
#export to path
$sitesData | Export-Csv -Path $csvFilePath -NoTypeInformation

# Check if the module is installed
if (-not (Get-Module -Name Mailozaurr -ListAvailable)) {
    Write-Host "Mailozaurr module is not installed. Attempting to install..."
    Install-Module -Name Mailozaurr -AllowClobber -Force
    if ($?) {
        Write-Host "Mailozaurr module installed successfully."
        Import-Module -Name Mailozaurr -Force
        Write-Host "Mailozaurr module imported."
    } else {
        Write-Host "Failed to install Mailozaurr module. Please check for errors."
        exit
    }
} else {
    Write-Host "Mailozaurr module is already installed."
    Import-Module -Name Mailozaurr -Force
    Write-Host "Mailozaurr module imported."
}

# Update the module
Write-Host "Checking for updates to Mailozaurr module..."
Update-Module -Name Mailozaurr
if ($?) {
    Write-Host "Mailozaurr module is up to date."
} else {
    Write-Host "Failed to update Mailozaurr module. Please check for errors."
}

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
    -Attachments $attachment `
    
