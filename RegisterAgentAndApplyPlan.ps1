[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, HelpMessage = "URL of the data center. Example: https://eu8.acronis.cloud")]
    [string]$DataCenterUrl,
    [Parameter(Mandatory = $false, HelpMessage = "Registration Token issued by the customer tenant. Example: ABCD-1234-5678.")]
    [string]$RegToken,
    [Parameter(Mandatory = $false, HelpMessage = "Boolean flag to specify if installation is to be performed. Default is false.")]
    [bool]$Install = $false,
    [Parameter(Mandatory = $false, HelpMessage = "Client ID for authentication. Not required.")]
    [string]$ClientId,
    [Parameter(Mandatory = $false, HelpMessage = "Client secret for authentication. Not required.")]
    [string]$ClientSecret,
    [Parameter(Mandatory = $false, HelpMessage = "ID of the customer tenant. Can be a UUIDv4 or an integer between 1000 and 10000000.")]
    [ValidateScript({
            if ($_ -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$' -or ($_ -is [int] -and $_ -ge 1000 -and $_ -le 10000000)) {
                $true
            }
            else {
                throw "CustomerTenantId must be either a UUIDv4 or an integer between 1000 and 10000000."
            }
        })]
    [string]$CustomerTenantId,
    [Parameter(Mandatory = $false, HelpMessage = "Policy ID in UUIDv4 format.")]
    [ValidateScript({
            if ($_ -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                $true
            }
            else {
                throw "PolicyId must be a valid UUIDv4."
            }
        })]
    [string]$PolicyId,
    [Parameter(Mandatory = $false, HelpMessage = "Personal Tenant ID in UUIDv4 format.")]
    [ValidateScript({
            if ($_ -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                $true
            }
            else {
                throw "PersonalTenantId must be a valid UUIDv4."
            }
        })]
    [string]$PersonalTenantId,
    [Parameter(Mandatory = $false, HelpMessage = "Path to the installer file. The path should end with '.exe'.")]
    [string]$InstallerPath,
    [switch]$Help
)

function Show-Help {
    Write-Host "Usage: .\RegisterAndApplyPlan.ps1 [-DataCenterUrl] <string> [-RegToken] <string> [-Install] <boolean> [-ClientId] <string> [-ClientSecret] <string> [-CustomerTenantId] <string> [-PolicyId] <string> [-InstallerPath] <string> [-PersonalTenantId] <string>"
    Write-Host "  -DataCenterUrl       URL of the data center. Example: https://dc1.acronis.cloud"
    Write-Host "  -RegToken            Registration Token issued by the customer tenant. Example: ABCD-1234-5678."
    Write-Host "  -Install             Boolean flag (0 OR 1) to specify if installation is to be performed. Default is false (0)."
    Write-Host "  -ClientId            Client ID for authentication. Not required."
    Write-Host "  -ClientSecret        Client secret for authentication. Not required."
    Write-Host "  -CustomerTenantId    ID of the customer tenant. Can be a UUIDv4 or an integer between 1000 and 10000000."
    Write-Host "  -PolicyId            Policy ID in UUIDv4 format."
    Write-Host "  -InstallerPath       Path to the installer file. The path should end with '.exe'."
    Write-Host "  -PersonalTenantId    Personal Tenant ID in UUIDv4 format."
    exit
}
<#
The script is used to register an agent and apply a plan to it.
Usage: .\RegisterAndApplyPlan.ps1 [-DataCenterUrl] <string> [-RegToken] <string> [-Install] <boolean> [-ClientId] <string> [-ClientSecret] <string> [-CustomerTenantId] <string> [-PolicyId] <string> [-InstallerPath] <string>"
DataCenterUrl       URL of the data center. Example: https://dc1.acronis.cloud
RegToken            Registration Token issued by the customer tenant. Example: ABCD-1234-5678.
Install             Boolean flag to specify if installation is to be performed. Default is false. 
ClientId            Client ID for authentication. Not required.
ClientSecret        Client secret for authentication. Not required.
CustomerTenantId    ID of the customer tenant. Can be a UUIDv4 or an integer between 1000 and 10000000.
PolicyId            Policy ID in UUIDv4 format.
InstallerPath       Path to the installer file. The path should end with '.exe'. If the path is not provided and the Install flag is set to true, 
the script will use the default installer path to Acronis Agent Installer for Windows. If the Install flag is false, it will use the default path to the RegisterAgentTool.
PersonalTenantId    Personal Tenant ID in UUIDv4 format.
#>

# Ask the user to enter the API Token Credentials and store them in the $client_id and $client_secret variables
function EnterApiTokenCredentials {

    param(
        [string]$logFilePath = $global:logFilePath
    )

    Add-Logfile -logFilePath $logFilePath -message "Asking user to enter the API Token Credentials..."

    # Helper function to check if a string input is null or only whitespace
    function IsNullOrWhiteSpace([string]$value) {
        return [string]::IsNullOrWhiteSpace($value)
    }

    # Function to get secure input or allow exit
    function GetSecureInput($prompt) {
        do {
            $input = Read-Host $prompt
            if ($input -eq 'exit') {
                Write-Host "Exiting..."
                exit
               # return $null
            }
            if ($input -eq 'default') {
                Write-Host "using default values..."
                return $false
            }
            if (IsNullOrWhiteSpace $input) {
                Write-Host "Input cannot be blank. Please enter a valid value or type 'exit' to abort."
                continue
            }
            try {
                $secureInput = ConvertTo-SecureString -String $input -AsPlainText -Force
            }
            catch {
                Write-Host "An error occurred converting input to secure string. Please try again."
                continue
            }
            return $secureInput
        } while ($true)
    }

    $client_id = GetSecureInput "Enter Client ID (or type 'exit' to abort)"
    if ($null -eq $client_id) {

        return [PSCustomObject]@{ client_id = $null; client_secret = $null }
    }
    
    $client_secret = GetSecureInput "Enter Client Secret (or type 'exit' to abort)"
    if ($null -eq $client_secret) {

        return [PSCustomObject]@{ client_id = $null; client_secret = $null }
    }

    return [PSCustomObject]@{
        client_id     = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($client_id))
        client_secret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($client_secret))
    }
}
function Send-AuthorizationRequest {
    param (
        [string]$client_id,
        [string]$client_secret,
        [string]$logFilePath = $global:logFilePath,
        [string]$dataCenterUrl
    )
    $unixRunTime, $utcRunTime, $formattedUtcRunTime, $filenameRunTime = Get-UnixAndUtcTimeStamp
    $responseOutput = Join-Path -Path $PSScriptRoot -ChildPath "Auth-Token-Response-$filenameRunTime.json"
    $headers = @{
        "Content-Type" = "application/x-www-form-urlencoded"
    }
    $body = @{
        "grant_type"    = "client_credentials"
        "client_id"     = $client_id
        "client_secret" = $client_secret
    }
    $uri = "$dataCenterUrl/api/2/idp/token"

    Add-Logfile -logFilePath $logFilePath -message "Sending authorization request to: $uri"
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body

        # Write the response to a file
        $response | ConvertTo-Json | Set-Content $responseOutput
        Add-Logfile -logFilePath $logFilePath -message "HTTP response: $response"
        # Return the response object
        return $response
    }
    catch {
        Add-Logfile -logFilePath $logFilePath -message "An error occurred sending the authorization request: $_"
        throw $_
    }
}
function Send-AuthorizationScopeRequest {
    param (
        [string]$tokenResponseAccessToken,
        [string]$targetTenantScope,
        [string]$logFilePath = $global:logFilePath,
        [string]$dataCenterUrl
    )
    $unixRunTime, $utcRunTime, $formattedUtcRunTime, $filenameRunTime = Get-UnixAndUtcTimeStamp
    $responseOutput = Join-Path -Path $PSScriptRoot -ChildPath "Scope-Auth-Token-Response-$filenameRunTime.json"
    $headers = @{
        "Content-Type" = "application/x-www-form-urlencoded"
    }
    $body = @{
        "grant_type" = "urn:ietf:params:oauth:grant-type:jwt-bearer"
        "assertion"  = $tokenResponseAccessToken
        "scope"      = "urn:acronis.com:tenant-id:" + $targetTenantScope
    }
    $uri = "$dataCenterUrl/bc/idp/token"

    Add-Logfile -logFilePath $logFilePath -message "Sending authorization request to: $uri"
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body

        # Write the response to a file
        $response | ConvertTo-Json | Set-Content $responseOutput
        Add-Logfile -logFilePath $logFilePath -message "Scoped authorization successful. Token received. HTTP response: $response"
        Write-Host "Scoped authorization successful. Token received."

        # Return the response object
        return $response
    }
    catch {
        Add-Logfile -logFilePath $logFilePath -message "An error occurred sending the authorization request: $_"
        throw $_
    }
}
function Send-RegTokenAuthorizationScopeRequest {
    param (
        [string]$regToken,
        [string]$logFilePath = $global:logFilePath,
        [string]$dataCenterUrl
    )
    $unixRunTime, $utcRunTime, $formattedUtcRunTime, $filenameRunTime = Get-UnixAndUtcTimeStamp
    $responseOutput = Join-Path -Path $PSScriptRoot -ChildPath "Reg-Token-Auth-Token-Response-$filenameRunTime.json"
    $headers = @{
        "Content-Type" = "application/x-www-form-urlencoded"
    }
    $body = @{
        "grant_type" = "urn:ietf:params:oauth:grant-type:jwt-bearer"
        "assertion"  = $regToken
    }
    $uri = "$dataCenterUrl/api/2/idp/token"

    Add-Logfile -logFilePath $logFilePath -message "Sending authorization request to: $uri"
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body

        # Write the response to a file
        $response | ConvertTo-Json | Set-Content $responseOutput
        Add-Logfile -logFilePath $logFilePath -message "Registration token scoped authorization successful. Token received. HTTP response: $response"
        Write-Host "Registration token scoped authorization successful. Token received."
        # Return the response object
        return $response
    }
    catch {
        Add-Logfile -logFilePath $logFilePath -message "An error occurred sending the authorization request: $_"
        throw $_
    }
}
function HandleAuthorization {
    param (
        [string]$client_id,
        [string]$client_secret,
        [string]$logFilePath = $global:logFilePath,
        [string]$dataCenterUrl
    )
    try {
        $response = Send-AuthorizationRequest -client_id $client_id -client_secret $client_secret -logFilePath $logFilePath -dataCenterUrl $dataCenterUrl
        Add-Logfile -logFilePath $logFilePath -message "Credentialed authorization successful. Token received. HTTP response: $response"
        Write-Host "Credentialed authorization successful. Token received."
        return $response
    }
    catch {
        $errorCode = $_.Exception.Response.StatusCode
        $errorDescription = $_.Exception.Response.StatusDescription
        Add-Logfile -logFilePath $logFilePath -message "Failed to authorize: $errorCode, $errorDescription"

        # Log this error as needed
        return $null
    }
}
function Get-UUIDFromInput {
    param (
        [string]$inputString
    )

    # Regular expression pattern to match UUID v4
    $uuidPattern = "[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-4[0-9a-fA-F]{3}-[89abAB][0-9a-fA-F]{3}-[0-9a-fA-F]{12}"
    # Regular expression pattern to find integers greater than 5 digits
    $intPattern = "\b\d{6,}\b"

    # Find matches for UUIDs and integers
    $uuidMatches = [regex]::Matches($inputString, $uuidPattern)
    $intMatches = [regex]::Matches($inputString, $intPattern)

    # Output findings and prompt user for selection if multiple matches are found
    if ($uuidMatches.Count -eq 0 -and $intMatches.Count -eq 0) {
        return $null
    }
    elseif ($uuidMatches.Count -eq 1 -and $intMatches.Count -eq 0) {
        return $uuidMatches[0].Value
    }
    elseif ($uuidMatches.Count -eq 0 -and $intMatches.Count -eq 1) {
        Write-Host "Found one integer greater than 5 digits: $($intMatches[0].Value)"
    }
    else {
        Write-Host "Multiple interesting items found. Please make a selection:"
        $index = 1

        # Display UUIDs with indexing
        foreach ($uuid in $uuidMatches) {
            Write-Host "$index : UUID - $($uuid.Value)"
            $index++
        }

        # Display integers with indexing
        foreach ($int in $intMatches) {
            Write-Host "$index : Integer - $($int.Value)"
            $index++
        }

        Write-Host "Enter the number to select an item or type 'exit' to exit."
        do {
            $input = Read-Host "Select a number or type 'exit'"
            if ($input -ieq 'exit') {
                Write-Host "Exiting..."
                exit
            }

            $selectedIndex = $input -as [int]
            $selectedIndex--

            if ($selectedIndex -lt 0 -or $selectedIndex -ge ($uuidMatches.Count + $intMatches.Count)) {
                Write-Host "Invalid selection, try again."
                $input = $null
            }
            elseif ($selectedIndex -lt $uuidMatches.Count) {
                return $uuidMatches[$selectedIndex].Value
            }
            else {
                $selectedIndex -= $uuidMatches.Count
                Write-Host "Selected integer greater than 5 digits: $($intMatches[$selectedIndex].Value)"
                return $intMatches[$selectedIndex].Value
            }
        } while ($input -eq $null)

    }
}
function Add-Logfile {
    param (
        [string]$logFilePath = $global:logFilePath,
        [string]$formattedTime = ((Get-Date).ToUniversalTime()).ToString("yyyy-MM-dd HH:mm:ss"),
        [string]$message
    )
    # Append log file message
    "$formattedTime - $message" | Out-File -FilePath $logFilePath -Append
}
function Write-Logfile {
    param (
        [string]$filenameStartTime,
        [string]$utcStartTime
    )

    # Define the log file name and path
    $logFilePath = Join-Path -Path $PSScriptRoot -ChildPath "Register-Agent-Apply-Plan-$filenameStartTime.log"
    
    # Create log file starting message
    "Log started at: $utcStartTime UTC time" | Out-File -FilePath $logFilePath -Append

    return $logFilePath
}
function Get-UnixAndUtcTimeStamp {
    # Get the current date and time in UTC
    $currentDate = (Get-Date).ToUniversalTime()
    # Format the current date and time in UTC to "yyyy-MM-dd HH:mm:ss"
    $formattedDate = $currentDate.ToString("yyyy-MM-dd HH:mm:ss")
    # Format the current date and time in UTC to "yyyyMMddHHmmss"
    $filenameDate = $currentDate.ToString("yyyyMMddHHmmss")
    # Create a datetime object for the Unix epoch start (January 1, 1970, at 00:00:00)
    $epochStart = [datetime]::new(1970, 1, 1, 0, 0, 0, [System.DateTimeKind]::Utc)
    # Calculate the total seconds between the Unix epoch start and the current UTC time
    $unixTime = [int](($currentDate - $epochStart).TotalSeconds)
    # Output the Unix time
    return $unixTime, $currentDate, $formattedDate, $filenameDate
}
function Read-Credential-Input {
    param (
        [string]$client_id,
        [string]$client_secret
    )

    if ($null -eq $client_id -or $null -eq $client_secret) {
        Add-Logfile -logFilePath $logFilePath -message "User aborted the script..."
        return $null
    }

    return $true
}
function Get-UserIds {
    param (
        [string]$tokenResponseAccessToken,
        [string]$tenantId,
        [string]$logFilePath = $global:logFilePath,
        [string]$dataCenterUrl
    )
    $unixRunTime, $utcRunTime, $formattedUtcRunTime, $filenameRunTime = Get-UnixAndUtcTimeStamp
    $responseOutput = Join-Path -Path $PSScriptRoot -ChildPath "Users-Subtree-Call-$filenameRunTime.json"
    $uri = "$dataCenterUrl/api/2/users?subtree_root_tenant_id=$tenantId"
    $headers = @{
        'Content-Type'  = "application/json"
        "Authorization" = "Bearer $($tokenResponse.access_token)"
    }
    Add-Logfile -logFilePath $logFilePath -message "Sending request to: $uri"
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
        # Write the response to a file
        $response | ConvertTo-Json | Set-Content $responseOutput
        Add-Logfile -logFilePath $logFilePath -message "HTTP response: $response"
        # Return the response object
        return $response
    }
    catch {
        Add-Logfile -logFilePath $logFilePath -message "An error occurred sending the request: $_"
        throw $_
    }
}
function New-RegistrationToken {
    param (
        [string]$tokenResponseAccessToken,
        [string]$customerTenantId,
        [string]$personalTenantId,
        [string]$policyId,
        [string]$logFilePath = $global:logFilePath,
        [string]$dataCenterUrl
    )
    $unixRunTime, $utcRunTime, $formattedUtcRunTime, $filenameRunTime = Get-UnixAndUtcTimeStamp
    $responseOutput = Join-Path -Path $PSScriptRoot -ChildPath "Registration-Token-Response-$filenameRunTime.json"
    $uri = "$dataCenterUrl/api/2/tenants/$customerTenantId/registration_tokens"
    $headers = @{
        'Content-Type'  = "application/json"
        "Authorization" = "Bearer $($tokenResponseAccessToken)"
    }
    [array]$scopes = @( "urn:acronis.com:tenant-id:$($personalTenantId):backup_agent_admin",
        "urn:acronis.com::policy_management:$($customerTenantId)|$($policyId):apply_revoke" )
    $body = @{
        "expires_in" = 3600
        "scopes"     = $scopes
    }
    Add-Logfile -logFilePath $logFilePath -message "Sending request to: $uri"
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body ($body | ConvertTo-Json -Depth 5)
        # Write the response to a file
        $response | ConvertTo-Json | Set-Content $responseOutput
        Add-Logfile -logFilePath $logFilePath -message "HTTP response: $response"
        # Return the response object
        return $response
    }
    catch {
        Add-Logfile -logFilePath $logFilePath -message "An error occurred sending the request: $_"
        throw $_
    }
}
function Get-ApplyPlantoResource {
    param (
        [string]$tokenResponseAccessToken,
        [string]$resourceId,
        [string]$logFilePath = $global:logFilePath,
        [string]$dataCenterUrl
    )

    $unixRunTime, $utcRunTime, $formattedUtcRunTime, $filenameRunTime = Get-UnixAndUtcTimeStamp
    $responseOutput = Join-Path -Path $PSScriptRoot -ChildPath "Apply-Plan-$filenameRunTime.json"

    $uri = "$dataCenterUrl/api/policy_management/v4/applications"
    $json = @"
{
    "accepted_default_name": "",

	 "context": {

		 "items": [

			 "${resourceId}"

		 ]

	 }

}
"@
    $headers = @{
        'Content-Type'  = "application/json"
        "Authorization" = "Bearer $($tokenResponseAccessToken)"
    }
    Add-Logfile -logFilePath $logFilePath -message "Sending request to: $uri"
    $timeout = New-TimeSpan -Minutes 10
    $endTime = (Get-Date).Add($timeout)
    $attempt = 0
    $totalAttempts = 60

    while ((Get-Date) -lt $endTime) {
        $attempt++
        Write-Progress -Activity "Attempting to register agent on device" -Status "Attempt: $attempt" -PercentComplete (($attempt / $totalAttempts) * 100)
        try {
            $response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $json
            # Write the response to a file
            $response | ConvertTo-Json | Set-Content $responseOutput
            Add-Logfile -logFilePath $logFilePath -message "Registration Successful: $response"
            # Return the response object
        }
        catch {
            if ($attempt -eq $totalAttempts) {
                Add-Logfile -logFilePath $logFilePath -message "An error occurred sending the request: $_"
                Write-Host "Resource: $($resourceId) not found in cloud."
            }
        }
        if ($null -ne $response) {
            Write-Host "Resource ID found: $($responseOutput)"
            return $response
        } 
        Start-Sleep -Seconds 10
    }
}

function Get-UserFromInput {
    param (
        [array]$userObjects
    )
    if ($userObjects.Count -eq 0) {
        Write-Host "No users found."
        return $null
    }
    
    #Write-Host $promptMessage
    for ($i = 0; $i -lt $userObjects.Count; $i++) {
        Write-Host "$($i+1): $($userObjects[$i].login) - $($userObjects[$i].personal_tenant_id)"
    }
    
    do {
        $selected = Read-Host "Enter the number of the user"
        if ($selected -lt 1 -or $selected -gt $userObjects.Count) {
            Write-Host "Invalid selection. Please select a number between 1 and $($userObjects.Count)."
        }
    } while ($selected -lt 1 -or $selected -gt $userObjects.Count)
    
    return $userObjects[$selected - 1]
}

function Find-ResourceIdPollRegistry {
    param (
        [Parameter(Mandatory = $true)]
        [string]$keyPath
    )

    $timeout = New-TimeSpan -Minutes 5
    $endTime = (Get-Date).Add($timeout)
    $attempt = 0
    $keyValue = $null

    while ((Get-Date) -lt $endTime) {
        $attempt++
        Write-Progress -Activity "Polling registry" -Status "Attempt: $attempt" -PercentComplete (($attempt / 30) * 100)
        $keyValue = Get-ItemProperty -Path $keyPath -ErrorAction SilentlyContinue 
        if ($null -ne $keyValue -and $null -ne $keyValue.InstanceID) {
            break
        }
        
        Start-Sleep -Seconds 10
    }
    
    if ($null -eq $keyValue) {
        Write-Host "RegistryId not found within the specified time."
    }
    else {
        Write-Host "Object found in registry: $keyValue"
        $keyValue = Get-ItemProperty -Path $keyPath -ErrorAction SilentlyContinue | Select-Object -ExpandProperty InstanceID
        Write-Host "Resource ID found in registry: $keyValue"
    }

    return $keyValue
}

function Get-ResourceById {
    param (
        [string]$tokenResponseAccessToken,
        [string]$resourceId,
        [string]$logFilePath = $global:logFilePath,
        [string]$dataCenterUrl
    )

    $unixRunTime, $utcRunTime, $formattedUtcRunTime, $filenameRunTime = Get-UnixAndUtcTimeStamp
    $responseOutput = Join-Path -Path $PSScriptRoot -ChildPath "Resource-Get-by-Id-$filenameRunTime.json"

    $uri = "$dataCenterUrl/api/resource_management/v4/resources/$resourceId"
    $headers = @{
        "Authorization" = "Bearer $($tokenResponseAccessToken)"
    }
    Add-Logfile -logFilePath $logFilePath -message "Sending request to: $uri"
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
        # Write the response to a file
        $response | ConvertTo-Json | Set-Content $responseOutput

        # Return the response object
        return $response
    }
    catch {
       # $errorDescription = $_.Exception.Response.StatusDescription
        Add-Logfile -logFilePath $logFilePath -message "An error occurred sending the request: $_"
        throw $_
    }
}
function Find-ResourceIdPollCloud {
    param (
        [Parameter(Mandatory = $true)]
        [string]$tokenResponseAccessToken,
        [string]$logFilePath = $global:logFilePath,
        [string]$resourceId,
        [string]$dataCenterUrl
    )

    $timeout = New-TimeSpan -Minutes 20
    $endTime = (Get-Date).Add($timeout)
    $attempt = 0
    $totalAttempts = 120

    while ((Get-Date) -lt $endTime) {
        $attempt++
        Write-Progress -Activity "Polling Cloud API for Resource" -Status "Attempt: $attempt" -PercentComplete (($attempt / $totalAttempts) * 100)
        try {
            $response = Get-ResourceById -tokenResponseAccessToken $tokenResponseAccessToken -resourceId $resourceId -logFilePath $logFilePath -dataCenterUrl $dataCenterUrl
            Write-Host $response | ConvertTo-Json
        }
        catch {
            if ($attempt -eq $totalAttempts) {
                Write-Host "An error occurred getting the resource by ID: $_"
                Add-Logfile -logFilePath $logFilePath -message "An error occurred getting the resource by ID: $_"
            }
        }
        if ($null -ne $response) {
            Write-Host "Resource ID found: $($response.id)"
            break
        }

        Start-Sleep -Seconds 10
    }

    if ($null -eq $response) {
        Write-Host "Resource ID: $resourceId not found within the specified time."
        Add-Logfile -logFilePath $logFilePath -message "Resource ID: $resourceId not found within the specified time."
        exit
    }
    else {
        $response = $($response.id)
    }

    return $keyValue
}

function Get-AgentInstallerUrl {
    param([string]$dataCenterUrl)
    $agentListUrl = "${dataCenterUrl}/bc/api/ams/links/list"
    $agentListResponse = Invoke-RestMethod -Uri $agentListUrl -Method Get
    foreach ($agent in $agentListResponse.auto_update) {
        if ($agent.system -eq "windows" -and $agent.architecture -eq (32, 64)[[System.Environment]::Is64BitOperatingSystem]) {
            $agentUrl = $agent.url
            Write-Host "Agent URL: $agentUrl"
            $agentUrl = $agentUrl -replace ("CyberProtect_Agent_web.exe", "CyberProtect_AgentForWindows_x64.exe") 
            $agentUrl = $agentUrl -replace ("dl.managed-protection.com", "dl.acronis.com") 
            Write-Host "Modified Agent URL: $agentUrl"
            return $agent.url
        }
    }
}

function Get-AgentFromUrl {
    param([string]$installerUrl, [string]$installerPath)
    $installerPath = Join-Path -Path $PSScriptRoot -ChildPath "AcronisCyberProtect_AgentForWindows_web.exe"
    $webClient = New-Object TimeoutWebClient
    try {
        $webClient.DownloadFile($installerUrl, $installerPath)
    }
    catch {
        Write-Error $Error[0]
        throw "Failure to install agent installer"
    }
    finally {
        $webClient.Dispose()
    }
    
}

if ($Help) {
    Show-Help
}

$timeoutWebClientCode = @"
public class TimeoutWebClient : System.Net.WebClient
{
    protected override System.Net.WebRequest GetWebRequest(System.Uri address)
    {
        System.Net.WebRequest request = base.GetWebRequest(address);

        if (request != null)
        {
            request.Timeout = System.Convert.ToInt32(System.TimeSpan.FromHours(1).TotalMilliseconds);
        }
        return request;
    }
}
"@;
Add-Type -TypeDefinition $timeoutWebclientCode -Language CSharp -WarningAction SilentlyContinue
$unixStartTime, $utcStartTime, $formattedUtcStartTime, $filenameStartTime = Get-UnixAndUtcTimeStamp

$global:logFilePath = Write-Logfile -filenameStartTime $filenameStartTime -utcStartTime $utcStartTime


# Using PSBoundParameters to check if parameters were provided
if ($PSBoundParameters.ContainsKey('DataCenterUrl')) {
    Write-Host "DataCenter URL was provided: $DataCenterUrl"
}
else {
    $DataCenterUrl = Read-Host "Enter the URL of the data center. Example: https://eu8.acronis.cloud"
    if ($null -eq $DataCenterUrl) {
        Write-Host "No data center URL provided... Exiting..."
        Add-Logfile -logFilePath $logFilePath -message "No data center URL provided... Exiting..."
        return
    }
}
if ($PSBoundParameters.ContainsKey('RegToken')) {
    Write-Host "Registration Token was provided: $RegToken"
    $regToken = $RegToken
    if ($Install) {
        Write-Host "Installing Acronis Agent..."
        $exeLogFilePath = Join-Path -Path $PSScriptRoot -ChildPath "ACPRegistrationCommand-$filenameStartTime.log"
        Write-Host "Log file path: $exeLogFilePath"
        if ($PSBoundParameters.ContainsKey('InstallerPath')) {
            $installerPath = $InstallerPath
        }
        else {
            $installerPath = Join-Path -Path $PSScriptRoot -ChildPath "AcronisCyberProtect_AgentForWindows_web.exe"
            if (-not (Test-Path -Path $installerPath)) {
                # If the agent installer is not in the script directory, download it
                $installerUrl = Get-AgentInstallerUrl -dataCenterUrl $DataCenterUrl
                Get-AgentFromUrl -installerUrl $installerUrl -installerPath $installerPath
            }
            $InstallerPath = $installerPath
        }
        [System.Diagnostics.Process]::Start("$($installerPath)", "--quiet --registration by-token --reg-token $($regToken) --reg-address $($DataCenterUrl) --log-dir $($exeLogFilePath) --add-components commandLine,agentForWindows,trayMonitor")
    }
    else {
        if ($PSBoundParameters.ContainsKey('InstallerPath')) {
            $installerPath = $InstallerPath
        } else {
            $installerPath = $env:ProgramFiles + "\BackupClient\RegisterAgentTool\register_agent.exe"
        }
        Write-Host "Acronis Agent will not be installed."
        [System.Diagnostics.Process]::Start("$($installerPath)", "-o register -t cloud -a $DataCenterUrl --token $($regToken)")
    }
    $keyPath = "HKLM:\SOFTWARE\Acronis\BackupAndRecovery\Settings\MachineManager"
    $keyValue = Find-ResourceIdPollRegistry -keyPath $keyPath
    
    $regTokenAuthResponse = Send-RegTokenAuthorizationScopeRequest -regToken $regToken -logFilePath $logFilePath -dataCenterUrl $DataCenterUrl
    Write-Host "Attempting to apply policy..."
    $result = Get-ApplyPlantoResource -tokenResponseAccessToken $regTokenAuthResponse.access_token -resourceId $keyValue -logFilePath $logFilePath -dataCenterUrl $DataCenterUrl
}
else {
    if ($PSBoundParameters.ContainsKey('ClientId') -and $PSBoundParameters.ContainsKey('ClientSecret')) {
        Write-Host "Client ID and Secret were provided: $ClientId"
    
        $tokenResponse = HandleAuthorization -client_id $ClientId -client_secret $ClientSecret -logFilePath $logFilePath -dataCenterUrl $DataCenterUrl
        if ($null -eq $tokenResponse) {
            Write-Host "Authorization failed. Please try again."
            do {
                if ($null -eq $tokenResponse) {
                    Write-Host "Authorization failed. Please try again."
                    $userCredentials = EnterApiTokenCredentials -logFilePath $logFilePath
        
                    $readCredentials = Read-Credential-Input -client_id $userCredentials.client_id -client_secret $userCredentials.client_secret
                    if ($null -eq $readCredentials) {
                        return
                    }
                    $tokenResponse = HandleAuthorization -client_id $userCredentials.client_id -client_secret $userCredentials.client_secret -logFilePath $logFilePath -dataCenterUrl $DataCenterUrl
                    if ($null -eq $tokenResponse) {
                        Write-Host "Authorization failed. Please try again."
                    }
                }
            } while ($null -eq $tokenResponse)
        } else {
            Write-Host "Authorization from provided Client ID and Secret successful."
        }
    
    }
    else {
        do {
            $userCredentials = EnterApiTokenCredentials -logFilePath $logFilePath
    
            $readCredentials = Read-Credential-Input -client_id $userCredentials.client_id -client_secret $userCredentials.client_secret
            if ($null -eq $readCredentials) {
                return
            }
            $tokenResponse = HandleAuthorization -client_id $userCredentials.client_id -client_secret $userCredentials.client_secret -logFilePath $logFilePath -dataCenterUrl $DataCenterUrl
            if ($null -eq $tokenResponse) {
                Write-Host "Authorization failed. Please try again."
            }
        } while ($null -eq $tokenResponse)
    }
    
    if ($PSBoundParameters.ContainsKey('PolicyId')) {
        Write-Host "Policy ID was provided: $PolicyId"
        $policyId = $PolicyId
    }
    else {
        $policyId = Read-Host "Enter the UUID of the policy you wish to apply"
        if ($null -eq $policyId) {
            Write-Host "No Policy ID provided... Exiting..."
            Add-Logfile -logFilePath $logFilePath -message "No Policy ID provided... Exiting..."
            return
        }
    }

    if ($PSBoundParameters.ContainsKey('Install')) {
        Write-Host "Install flag was provided: $Install"
    }
    else {
        Write-Host "Install flag was not provided. Default is false."
    }
    if ($PSBoundParameters.ContainsKey('CustomerTenantId')) {
        Write-Host "Customer Tenant ID was provided: $CustomerTenantId"
        $targetTenantScope = $CustomerTenantId
    }
    else {
        $authenticatedUuid = Get-UUIDFromInput -inputString $tokenResponse.scope
        Write-Host "Customer Tenant ID was not provided. Using authenticated UUID: $authenticatedUuid"
        if ($null -eq $authenticatedUuid) {
            Write-Host "No UUID found... Exiting..."
            Add-Logfile -logFilePath $logFilePath -message "No UUID found... Exiting..."
            return
        }
        # Ask the user to enter the UUID of the tenant they wish to scope into
        $targetTenantScope = $authenticatedUuid
    }
    
    $scopedResponse = Send-AuthorizationScopeRequest -tokenResponseAccessToken $tokenResponse.access_token -targetTenantScope $targetTenantScope -logFilePath $logFilePath -dataCenterUrl $DataCenterUrl
    if ($PSBoundParameters.ContainsKey('PersonalTenantId')) {
        Write-Host "Personal Tenant ID was provided: $PersonalTenantId"
        $personalTenantId = $PersonalTenantId
        $newRegistrationToken = New-RegistrationToken -tokenResponseAccessToken $scopedResponse.access_token -customerTenantId $targetTenantScope -personalTenantId $PersonalTenantId -logFilePath $logFilePath -dataCenterUrl $DataCenterUrl -policyId $policyId
        
    }
    else {
        Write-Host "Personal Tenant ID was not provided."
        $usersResponse = Get-UserIds -tokenResponseAccessToken $scopedResponse.access_token -tenantId $targetTenantScope -logFilePath $logFilePath -dataCenterUrl $DataCenterUrl
        $selectedUser = Get-UserFromInput -userObjects $usersResponse.items -logFilePath $logFilePath
        $newRegistrationToken = New-RegistrationToken -tokenResponseAccessToken $scopedResponse.access_token -customerTenantId $selectedUser.tenant_id -personalTenantId $selectedUser.personal_tenant_id -logFilePath $logFilePath -dataCenterUrl $DataCenterUrl -policyId $policyId
    }
    $regToken = $newRegistrationToken.token
    
    if ($Install) {
        Write-Host "Installing Acronis Agent..."
        $exeLogFilePath = Join-Path -Path $PSScriptRoot -ChildPath "ACPRegistrationCommand-$filenameStartTime.log"
        Write-Host "Log file path: $exeLogFilePath"
        if ($PSBoundParameters.ContainsKey('InstallerPath')) {
            $installerPath = $InstallerPath
        }
        else {
            # Check if the agent installer is in the script directory
            $installerPath = Join-Path -Path $PSScriptRoot -ChildPath "AcronisCyberProtect_AgentForWindows_web.exe"
            if (-not (Test-Path -Path $installerPath)) {
                # If the agent installer is not in the script directory, download it
                $installerUrl = Get-AgentInstallerUrl -dataCenterUrl $DataCenterUrl
                Get-AgentFromUrl -installerUrl $installerUrl -installerPath $installerPath
            }
            $InstallerPath = $installerPath
        }
        [System.Diagnostics.Process]::Start("$($installerPath)", "--quiet --registration by-token --reg-token $($regToken) --reg-address $($DataCenterUrl) --log-dir $($exeLogFilePath) --add-components commandLine,agentForWindows,trayMonitor")
    }
    else {
    
        if ($PSBoundParameters.ContainsKey('InstallerPath')) {
            $installerPath = $InstallerPath
        } else {
            $installerPath = $env:ProgramFiles + "\BackupClient\RegisterAgentTool\register_agent.exe"
        }
        Write-Host "Acronis Agent will not be installed."
        [System.Diagnostics.Process]::Start("$($installerPath)", "-o register -t cloud -a $DataCenterUrl --token $($regToken)")
    }
    $keyPath = "HKLM:\SOFTWARE\Acronis\BackupAndRecovery\Settings\MachineManager"
    $keyValue = Find-ResourceIdPollRegistry -keyPath $keyPath
    
    $regTokenAuthResponse = Send-RegTokenAuthorizationScopeRequest -regToken $regToken -logFilePath $logFilePath -dataCenterUrl $DataCenterUrl
    Write-Host "Checking for Resource ID in Cloud API..."
    $resourceId = Find-ResourceIdPollCloud -tokenResponseAccessToken $scopedResponse.access_token -logFilePath $logFilePath -resourceId $keyValue -dataCenterUrl $DataCenterUrl
    
    
    $result = Get-ApplyPlantoResource -tokenResponseAccessToken $regTokenAuthResponse.access_token -resourceId $keyValue -logFilePath $logFilePath -dataCenterUrl $DataCenterUrl
}


$unixEndTime, $utcEndTime, $formattedUtcEndTime, $filenameEndTime = Get-UnixAndUtcTimeStamp

Add-Logfile -logFilePath $logFilePath -message "Finished processing services in: ($($unixEndTime - $unixStartTime) seconds). Finished at $utcEndTime UTC time"

Write-Host "Finished processing services in: ($($unixEndTime - $unixStartTime) seconds). Log file written to: $logFilePath`n"
