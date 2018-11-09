# Mandatory Parameter - Webex site name
Param([Parameter(Mandatory=$true)][string]$site)

# Function to write output to a log file
Function Write-Log {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")]
        [String]
        $Level = "INFO",

        [Parameter(Mandatory=$True)]
        [string]
        $Message,

        [Parameter(Mandatory=$False)]
        [string]
        $logfile
    )

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp $Level $Message"
    If($logfile) {
        Add-Content $logfile -Value $Line
    }
    Else {
        Write-Output $Line
    }
}

# Create log file
$today = (Get-Date).toString("yyyy-MM-dd")
$logFile = "$PSScriptRoot\{0}_WebexDownload.log" -f $today

Write-Log "INFO" "----------- Webex Client Desktop App Download: Starting ----------" $logFile

# Specify supported security types
$AllProtocols = [System.Net.SecurityProtocolType]'Ssl3,Tls,Tls11,Tls12'
[System.Net.ServicePointManager]::SecurityProtocol = $AllProtocols

Write-Log "INFO" "Set security types" $logFile

# Define Webex Uri and send the web request
$webexUri = "https://{0}.webex.com/mw3300/mywebex/default.do?siteurl={0}&service=10" -f $site

Write-Log "INFO" "Set webex Uri as $webexUri" $logFile
Write-Log "INFO" "Calling Webex Uri now" $logFile
try {
    $WebResponse = Invoke-WebRequest -Uri $webexUri
} catch {
    $message = "Received error while calling Webex Uri: {0}" -f $_.Exception.Message
    Write-Log "ERROR" $message $logFile
    Write-Log "ERROR" "----------- Webex Client Desktop App Download: Completed (Failed) ----------" $logFile
    break
}
Write-Log "INFO" "Received valid response from Webex Uri" $logFile

# Extract the Webex desktop app client version from the Web response
Write-Log "INFO" "Performing pattern-matching of Webex desktop app version from Webex Uri response" $logFile
$response = $WebResponse.Content | Select-String -Pattern ".*(WBXclient\-33.*)/.*"

# Assign the client version from the pattern matches
Write-Log "INFO" "Trying to extract Webex client version" $logFile
if ($response.Matches.Groups) {
    Write-Log "INFO" "Pattern-matching successful" $logFile
} else {
    Write-Log "ERROR" "Pattern-matching failed.  Saving output to webResponse.txt " $logFile
    $WebResponse.Content > "webResponse.txt"
    break
}

Write-Log "INFO" "Extracting capture group 1 from pattern-matching" $logFile
$clientVersion = $response.Matches.Groups[1].Value

Write-Log "INFO" "Setting client version to $clientVersion" $logFile

# Define the target directory to save the MSI
$fileName = "$PSScriptRoot\{0}.msi" -f $clientVersion
Write-Log "INFO" "Setting output MSI to $fileName" $logFile

# Define the Webex CDN Uri and send the web request/store the file
$cdnUri = "https://akamaicdn.webex.com/client/{0}/webexapp.msi" -f $clientVersion
Write-Log "INFO" "Setting CDN Uri to $cdnUri" $logFile

Write-Log "INFO" "Calling CDN Uri now" $logFile
try {
    $WebResponse = Invoke-WebRequest -Uri $cdnUri -OutFile $fileName
} catch {
    $message = "Received error while calling CDN Uri: {0}" -f $_.Exception.Message
    Write-Log "ERROR" $message $logFile 
    Write-Log "ERROR" "----------- Webex Client Desktop App Download: Completed (Failed) ----------" $logFile
    break
}

Write-Log "INFO" "----------- Webex Client Desktop App Download: Completed (Success) ----------" $logFile
