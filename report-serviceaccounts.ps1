# Set variables for HTML output file
$outputFile = "C:\temp\serviceaccounts.html"
# Uncomment the below if you need to use credentials other than the logged in user
# Attempt will be made for logged in user first, so can use two separate accounts for full range of servers
#$cred = Get-Credential

# Get list of servers from Active Directory
$servers = Get-ADComputer -Filter {OperatingSystem -Like "Windows Server*"} -Property Name |
    Select-Object -ExpandProperty Name

# Initialize an empty array for error messages
$errors = @()

# Loop through servers and retrieve service information
$results = foreach ($server in $servers) {
    Write-Host "Processing server $server"
    $connected = $false
    while (-not $connected) {
        try {
            $services = Get-WmiObject -Class Win32_Service -ComputerName $server -ErrorAction Stop
            $connected = $true
        }
        catch {
            try {
                $services = Get-WmiObject -Class Win32_Service -ComputerName $server -Credential $cred -ErrorAction Stop
                $connected = $true
            }
            catch {
                $errorMsg = "Could not connect to $server"
                Write-Warning $errorMsg
                $errors += $errorMsg
            }
        }
    }
    foreach ($service in $services) {
        $account = $service.StartName
        $serviceName = $service.Name
        # Exclude results where account is a system account
        if (($account -notmatch "^((NT AUTHORITY)|(BUILTIN))\\") -and ($account -ne $null) -and ($account -notlike "*localsystem*")) {

            [PSCustomObject]@{
                ServerName = $server
                ServiceName = $serviceName
                Account = $account
            }
        }
    }
}

# Convert results and errors to HTML tables
$tableResults = $results | ConvertTo-Html -Property ServerName,ServiceName,Account -As Table -Fragment
$tableErrors = $errors | ConvertTo-Html -As Table -Fragment

# Combine tables and output to HTML file
$html = "<html><head><style>table { border-collapse: collapse; } th, td { border: 1px solid black; padding: 5px; }</style></head><body>"
$html += "<h1>Service Account Information</h1>"
$html += $tableResults
$html += "<h1>Errors</h1>"
$html += $tableErrors
$html += "</body></html>"
$html | Out-File $outputFile
