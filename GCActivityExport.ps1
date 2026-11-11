# Garmin Connect Activity Export
# Mark van Eijken 2024-2026
# Version 2.0
# Version history:
# 1.0   - Initial version
# 1.1   - Fix: Garmin now expects parameters in the SSO url
#       Update: Added settings support in separate XML files
# 1.2   - Update support for new Garmin activity feed
# 1.3   - Update to support the new Garmin Signin URL
# 1.3.1 - Fix due error 402 Payment required error when retrieving activity list
# 1.3.2 - Remove default destination folder, because it causes to ignore the GCUserSettings.xml destination
#         Remove empty strings for Username and Password
#         Format layout
# 1.4   - Update authentication process to use an OAuth token
#         Add support for KML files
#         Updated user-agent
#         Remove Garmin Connect Actvity Export -Program Settings.xml dependency
#         Layout and output improvements
#         Report when rate limited (or any other error from Garmin)
# 2.0   - Rewrite of the authentication: Garmin retired the CAS form login + connect.garmin.com/modern/di-oauth/exchange
#         token endpoint. The script now uses the Garmin SSO JSON login (mobile flow, portal flow as fallback),
#         exchanges the service ticket for a DI OAuth2 token at diauth.garmin.com and calls connectapi.garmin.com.
#         Support for MFA (multi-factor authentication) codes
#         Token cache with automatic refresh, so a login (and MFA prompt) is not needed on every run
#         Retry with back-off when Garmin rate limits (HTTP 429) API calls
#         Fix: GCUserSettings.xml was never loaded
#         Fix: errors were silently swallowed by the catch block
#         Fix: delta cookie is only updated when all downloads succeeded
#         Compatible with Windows PowerShell 5.1 and PowerShell 7+ (Windows, Linux, macOS)
# The scripts does the following:
# - Downloads activity files from garmin in FIT, TCX, GPX or KML format.
# - Supports delta download

[CmdletBinding()]
PARAM(
    [ValidateSet("FIT", "TCX", "GPX", "KML")]
    [string]$ActivityFileType,
    [ValidateSet("All", "New")]
    [string]$DownloadOption,
    [string]$Destination,
    [string]$Username,
    [string]$Password,
    [ValidateSet("Yes", "No")]
    [string]$Overwrite,
    # Path of the user settings XML file. Defaults to GCUserSettings.xml next to this script.
    [string]$UserSettingsFile,
    # Path of the token cache file. Defaults to a per-user location (see Get-GCDefaultTokenCachePath).
    [string]$TokenCacheFile,
    # Do not read or write the token cache; always perform a full login.
    [switch]$NoTokenCache
)

#region Constants
$script:GC = [PSCustomObject]@{
    SsoBase           = "https://sso.garmin.com"
    ConnectApiBase    = "https://connectapi.garmin.com"
    DiTokenUrl        = "https://diauth.garmin.com/di-oauth2-service/oauth/token"
    DiGrantType       = "https://connectapi.garmin.com/di-oauth2-service/oauth/grant/service_ticket"
    # DI client ids used by the Garmin Connect mobile app. Garmin rotates these; the first one that works is used.
    DiClientIds       = @(
        "GARMIN_CONNECT_MOBILE_ANDROID_DI_2025Q2",
        "GARMIN_CONNECT_MOBILE_ANDROID_DI_2024Q4",
        "GARMIN_CONNECT_MOBILE_ANDROID_DI",
        "GARMIN_CONNECT_MOBILE_IOS_DI"
    )
    # Login flows, tried in order. Both answer with the same JSON shape.
    LoginFlows        = @(
        [PSCustomObject]@{
            Name       = "mobile"
            ClientId   = "GCM_IOS_DARK"
            ServiceUrl = "https://mobile.integration.garmin.com/gcm/ios"
            SignInPath = "/mobile/sso/en_US/sign-in"
            ApiPath    = "/mobile/api"
            UserAgent  = "Mozilla/5.0 (iPhone; CPU iPhone OS 18_7 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148"
        },
        [PSCustomObject]@{
            Name       = "portal"
            ClientId   = "GarminConnect"
            ServiceUrl = "https://connect.garmin.com/app"
            SignInPath = "/portal/sso/en-US/sign-in"
            ApiPath    = "/portal/api"
            UserAgent  = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36"
        }
    )
    NativeHeaders     = @{
        "X-Garmin-User-Agent"         = "com.garmin.android.apps.connectmobile/5.23; ; Google/sdk_gphone64_arm64/google; Android/33; Dalvik/2.1.0"
        "X-Garmin-Paired-App-Version" = "10861"
        "X-Garmin-Client-Platform"    = "Android"
        "X-App-Ver"                   = "10861"
        "X-Lang"                      = "en"
        "X-GCExperience"              = "GC5"
        "Accept-Language"             = "en-US,en;q=0.9"
    }
    NativeUserAgent   = "GCM-Android-5.23"
    ActivitySearchPath = "/activitylist-service/activities/search/activities"
    DownloadPaths     = @{
        FIT = "/download-service/files/activity/"
        TCX = "/download-service/export/tcx/activity/"
        GPX = "/download-service/export/gpx/activity/"
        KML = "/download-service/export/kml/activity/"
    }
    PageSize          = 100
    MaxRetries        = 4
}
#endregion

#region Generic helpers
function Test-GCIsWindows {
    # $IsWindows does not exist in Windows PowerShell 5.1.
    if ($PSVersionTable.PSEdition -eq "Desktop") { return $true }
    return [bool]$IsWindows
}

function ConvertTo-GCPlainText {
    param([Parameter(Mandatory)][System.Security.SecureString]$SecureString)
    return [System.Net.NetworkCredential]::new("", $SecureString).Password
}

function Get-GCSafeFileName {
    <# Builds the export file name (without extension). Keeps the naming of version 1.x so existing exports are recognised. #>
    param(
        [Parameter(Mandatory)]$ActivityId,
        [Parameter(Mandatory)][datetime]$StartTimeLocal,
        [string]$ActivityType,
        [string]$ActivityName
    )
    $name = ("{0} - {1} - {2} - {3}" -f $ActivityId, $StartTimeLocal.ToString("yyyy-MM-dd"), $ActivityType, $ActivityName).TrimEnd()
    return ($name -replace '[^A-Za-z0-9-_\@\,\(\) \.\[\]]', '-')
}

function ConvertFrom-GCJwtPayload {
    <# Decodes the payload of a JWT without validating the signature (only used to read exp/client_id). #>
    param([Parameter(Mandatory)][string]$Token)
    $parts = $Token.Split(".")
    if ($parts.Count -lt 2) { return $null }
    $payload = $parts[1].Replace("-", "+").Replace("_", "/")
    switch ($payload.Length % 4) {
        2 { $payload += "==" }
        3 { $payload += "=" }
    }
    try {
        $json = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($payload))
        return ($json | ConvertFrom-Json)
    }
    catch { return $null }
}

function Get-GCTokenExpiry {
    <# Returns the expiry of a JWT as UTC DateTime, or $null when it cannot be determined. #>
    param([Parameter(Mandatory)][string]$Token)
    $payload = ConvertFrom-GCJwtPayload -Token $Token
    if ($null -eq $payload -or $null -eq $payload.exp) { return $null }
    return [DateTime]::new(1970, 1, 1, 0, 0, 0, [DateTimeKind]::Utc).AddSeconds([double]$payload.exp)
}

function ConvertTo-GCHeaderTable {
    <# Normalises the different header collection types of PowerShell 5.1 and 7 into a case-insensitive hashtable. #>
    param($Headers)
    $table = @{}
    if ($null -eq $Headers) { return $table }
    if ($Headers -is [System.Net.WebHeaderCollection]) {
        foreach ($key in $Headers.AllKeys) { $table[$key] = $Headers[$key] }
    }
    elseif ($Headers -is [System.Collections.IDictionary]) {
        foreach ($key in @($Headers.Keys)) { $table[[string]$key] = (@($Headers[$key]) -join ",") }
    }
    else {
        # System.Net.Http.Headers.HttpHeaders: enumerable of KeyValuePair<string, IEnumerable<string>>
        foreach ($pair in $Headers) { $table[[string]$pair.Key] = (@($pair.Value) -join ",") }
    }
    return $table
}

function Invoke-GCHttp {
    <#
    Thin wrapper around Invoke-WebRequest that never throws on HTTP error codes and behaves the same on
    Windows PowerShell 5.1 and PowerShell 7. Returns StatusCode, Content (string), Bytes and Headers (hashtable).
    Throws only on transport errors (DNS, TLS, connection reset).
    #>
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$Method = "GET",
        [hashtable]$Headers = @{},
        $Body,
        [string]$ContentType,
        [string]$UserAgent,
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession
    )
    $params = @{
        Uri             = $Uri
        Method          = $Method
        Headers         = $Headers
        UseBasicParsing = $true
        ErrorAction     = "Stop"
    }
    if ($PSBoundParameters.ContainsKey("Body")) { $params.Body = $Body }
    if ($ContentType) { $params.ContentType = $ContentType }
    if ($UserAgent) { $params.UserAgent = $UserAgent }
    if ($WebSession) { $params.WebSession = $WebSession }
    if ($PSVersionTable.PSVersion.Major -ge 7) { $params.SkipHttpErrorCheck = $true }

    try {
        $response = Invoke-WebRequest @params
        $bytes = $null
        if ($null -ne $response.RawContentStream) { $bytes = $response.RawContentStream.ToArray() }
        $content = $response.Content
        if ($content -is [byte[]]) { $content = [System.Text.Encoding]::UTF8.GetString($content) }
        return [PSCustomObject]@{
            StatusCode = [int]$response.StatusCode
            Content    = [string]$content
            Bytes      = $bytes
            Headers    = (ConvertTo-GCHeaderTable -Headers $response.Headers)
        }
    }
    catch {
        # Windows PowerShell 5.1 throws a WebException on HTTP error codes.
        $errorResponse = $null
        if ($_.Exception.PSObject.Properties["Response"]) { $errorResponse = $_.Exception.Response }
        if ($null -eq $errorResponse) { throw }

        $content = $null
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $content = $_.ErrorDetails.Message }
        elseif ($errorResponse -is [System.Net.WebResponse]) {
            try {
                $reader = [System.IO.StreamReader]::new($errorResponse.GetResponseStream())
                $content = $reader.ReadToEnd()
                $reader.Dispose()
            }
            catch { $content = $null }
        }
        return [PSCustomObject]@{
            StatusCode = [int]$errorResponse.StatusCode
            Content    = [string]$content
            Bytes      = $null
            Headers    = (ConvertTo-GCHeaderTable -Headers $errorResponse.Headers)
        }
    }
}

function ConvertFrom-GCJsonSafe {
    param([string]$Content)
    if ([string]::IsNullOrWhiteSpace($Content)) { return $null }
    try { return ($Content | ConvertFrom-Json -ErrorAction Stop) } catch { return $null }
}
#endregion

#region Authentication
function Get-GCDefaultTokenCachePath {
    if (Test-GCIsWindows) { $base = $env:LOCALAPPDATA }
    else {
        $base = $env:XDG_CONFIG_HOME
        if ([string]::IsNullOrEmpty($base)) { $base = Join-Path -Path $HOME -ChildPath ".config" }
    }
    return (Join-Path -Path (Join-Path -Path $base -ChildPath "GarminConnectActivityExport") -ChildPath "GCTokenCache.xml")
}

function Save-GCTokenCache {
    <#
    Stores the DI tokens. On Windows the refresh token is encrypted with DPAPI (current user).
    On Linux/macOS PowerShell cannot encrypt a SecureString without a key, so the file is restricted to the
    current user (chmod 600) instead.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$Token,
        [Parameter(Mandatory)][string]$Username
    )
    $folder = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $folder)) { $null = New-Item -Path $folder -ItemType Directory -Force }
    $toSecure = { param($value) ConvertTo-SecureString -String $value -AsPlainText -Force | ConvertFrom-SecureString }
    $cache = [PSCustomObject]@{
        Username     = $Username.ToLowerInvariant()
        ClientId     = $Token.ClientId
        AccessToken  = (& $toSecure $Token.AccessToken)
        RefreshToken = (& $toSecure $Token.RefreshToken)
    }
    if (-not (Test-GCIsWindows)) {
        if (-not (Test-Path -LiteralPath $Path)) { $null = New-Item -Path $Path -ItemType File -Force }
        try { [System.IO.File]::SetUnixFileMode($Path, [System.IO.UnixFileMode]"UserRead, UserWrite") }
        catch { Write-Warning "WARNING - Could not restrict permissions of token cache $Path. Use -NoTokenCache if this is a shared system." }
    }
    $cache | Export-Clixml -LiteralPath $Path -Force
}

function Read-GCTokenCache {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Username)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try {
        $cache = Import-Clixml -LiteralPath $Path
        if ($cache.Username -ne $Username.ToLowerInvariant()) {
            Write-Verbose "Token cache belongs to another user, ignoring it."
            return $null
        }
        return [PSCustomObject]@{
            ClientId     = $cache.ClientId
            AccessToken  = ConvertTo-GCPlainText -SecureString (ConvertTo-SecureString -String $cache.AccessToken)
            RefreshToken = ConvertTo-GCPlainText -SecureString (ConvertTo-SecureString -String $cache.RefreshToken)
        }
    }
    catch {
        Write-Warning "WARNING - Token cache $Path could not be read and is ignored: $($_.Exception.Message)"
        return $null
    }
}

function New-GCTokenObject {
    param([Parameter(Mandatory)]$TokenResponse, [Parameter(Mandatory)][string]$ClientId)
    $payload = ConvertFrom-GCJwtPayload -Token $TokenResponse.access_token
    if ($payload -and $payload.client_id) { $ClientId = $payload.client_id }
    return [PSCustomObject]@{
        AccessToken  = [string]$TokenResponse.access_token
        RefreshToken = [string]$TokenResponse.refresh_token
        ClientId     = $ClientId
    }
}

function Invoke-GCDiTokenRequest {
    param([Parameter(Mandatory)][string]$ClientId, [Parameter(Mandatory)][hashtable]$Form)
    $basic = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("${ClientId}:"))
    $headers = $script:GC.NativeHeaders.Clone()
    $headers["Authorization"] = "Basic $basic"
    $headers["Accept"] = "application/json"
    $headers["Cache-Control"] = "no-cache"
    $body = ($Form.GetEnumerator() | ForEach-Object { "{0}={1}" -f $_.Key, [System.Uri]::EscapeDataString([string]$_.Value) }) -join "&"
    return Invoke-GCHttp -Uri $script:GC.DiTokenUrl -Method POST -Headers $headers -UserAgent $script:GC.NativeUserAgent `
        -ContentType "application/x-www-form-urlencoded" -Body $body
}

function Get-GCDiToken {
    <# Exchanges a CAS service ticket for a DI OAuth2 access + refresh token. #>
    param([Parameter(Mandatory)][string]$ServiceTicket, [Parameter(Mandatory)][string]$ServiceUrl)
    $lastStatus = $null
    foreach ($clientId in $script:GC.DiClientIds) {
        $response = Invoke-GCDiTokenRequest -ClientId $clientId -Form @{
            client_id      = $clientId
            service_ticket = $ServiceTicket
            grant_type     = $script:GC.DiGrantType
            service_url    = $ServiceUrl
        }
        $lastStatus = $response.StatusCode
        if ($response.StatusCode -eq 429) { throw "Garmin rate limited the token exchange (HTTP 429). Please try again later." }
        if ($response.StatusCode -ne 200) {
            Write-Verbose ("DI token exchange with client id {0} failed: HTTP {1}" -f $clientId, $response.StatusCode)
            continue
        }
        $json = ConvertFrom-GCJsonSafe -Content $response.Content
        if ($json -and $json.access_token) { return (New-GCTokenObject -TokenResponse $json -ClientId $clientId) }
    }
    throw "Exchanging the Garmin service ticket for an access token failed for all known client ids (last HTTP status: $lastStatus)."
}

function Update-GCDiToken {
    <# Uses the refresh token to get a new access token. Returns $null when the refresh token is no longer valid. #>
    param([Parameter(Mandatory)]$Token)
    if ([string]::IsNullOrEmpty($Token.RefreshToken) -or [string]::IsNullOrEmpty($Token.ClientId)) { return $null }
    $response = Invoke-GCDiTokenRequest -ClientId $Token.ClientId -Form @{
        grant_type    = "refresh_token"
        client_id     = $Token.ClientId
        refresh_token = $Token.RefreshToken
    }
    if ($response.StatusCode -ne 200) {
        Write-Verbose ("Refreshing the access token failed: HTTP {0}" -f $response.StatusCode)
        return $null
    }
    $json = ConvertFrom-GCJsonSafe -Content $response.Content
    if (-not ($json -and $json.access_token)) { return $null }
    $newToken = New-GCTokenObject -TokenResponse $json -ClientId $Token.ClientId
    if ([string]::IsNullOrEmpty($newToken.RefreshToken)) { $newToken.RefreshToken = $Token.RefreshToken }
    return $newToken
}

function Resolve-GCLoginResponse {
    <#
    Interprets the JSON answer of /{mobile|portal}/api/login and /api/mfa/verifyCode.
    Returns an object with Status (Success, MfaRequired, InvalidCredentials, RateLimited, Blocked, Error),
    Ticket, MfaMethod and Message.
    #>
    param([Parameter(Mandatory)]$Response)
    $result = [PSCustomObject]@{ Status = "Error"; Ticket = $null; MfaMethod = $null; Message = $null }
    if ($Response.StatusCode -in 427, 429) {
        $result.Status = "RateLimited"; $result.Message = "HTTP $($Response.StatusCode)"; return $result
    }
    $json = ConvertFrom-GCJsonSafe -Content $Response.Content
    if ($null -eq $json) {
        if ($Response.StatusCode -eq 403) { $result.Status = "Blocked" }
        $result.Message = "HTTP $($Response.StatusCode), no JSON response (probably a Cloudflare challenge)"
        return $result
    }
    if ($json.error -and $json.error.'status-code' -eq "429") {
        $result.Status = "RateLimited"; $result.Message = "429 in response body"; return $result
    }
    $type = $null
    if ($json.responseStatus) { $type = $json.responseStatus.type }
    switch ($type) {
        "SUCCESSFUL" { $result.Status = "Success"; $result.Ticket = $json.serviceTicketId }
        "MFA_REQUIRED" {
            $result.Status = "MfaRequired"
            $result.MfaMethod = "email"
            if ($json.customerMfaInfo -and $json.customerMfaInfo.mfaLastMethodUsed) { $result.MfaMethod = $json.customerMfaInfo.mfaLastMethodUsed }
        }
        "INVALID_USERNAME_PASSWORD" { $result.Status = "InvalidCredentials" }
        "CAPTCHA_REQUIRED" { $result.Status = "Blocked"; $result.Message = "Garmin requires a CAPTCHA" }
        default { $result.Message = "HTTP $($Response.StatusCode), responseStatus=$type" }
    }
    return $result
}

function Invoke-GCSsoLogin {
    <#
    Logs in at Garmin SSO with one login flow and returns the service ticket.
    Throws with a message starting with RETRYABLE: when the next flow may succeed.
    #>
    param(
        [Parameter(Mandatory)]$Flow,
        [Parameter(Mandatory)][string]$Username,
        [Parameter(Mandatory)][System.Security.SecureString]$Password,
        [scriptblock]$MfaCodeProvider
    )
    $query = "clientId={0}&service={1}" -f $Flow.ClientId, [System.Uri]::EscapeDataString($Flow.ServiceUrl)
    $signInUrl = "{0}{1}?{2}" -f $script:GC.SsoBase, $Flow.SignInPath, $query
    $apiQuery = "clientId={0}&locale=en-US&service={1}" -f $Flow.ClientId, [System.Uri]::EscapeDataString($Flow.ServiceUrl)

    # Step 1: load the sign-in page like a browser would. It sets the SESSION and Cloudflare cookies;
    # a credential POST without them is answered with 429.
    $session = [Microsoft.PowerShell.Commands.WebRequestSession]::new()
    $page = Invoke-GCHttp -Uri $signInUrl -WebSession $session -UserAgent $Flow.UserAgent -Headers @{
        "Accept"          = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        "Accept-Language" = "en-US,en;q=0.9"
    }
    if ($page.StatusCode -in 427, 429) { throw "RETRYABLE: sign-in page rate limited (HTTP $($page.StatusCode))" }
    if ($page.StatusCode -ne 200) { throw "RETRYABLE: sign-in page returned HTTP $($page.StatusCode)" }

    # Garmin/Cloudflare flags an immediate GET -> POST sequence as bot traffic.
    Start-Sleep -Milliseconds (Get-Random -Minimum 2000 -Maximum 5000)

    $postHeaders = @{
        "Accept"          = "application/json, text/plain, */*"
        "Accept-Language" = "en-US,en;q=0.9"
        "Origin"          = $script:GC.SsoBase
        "Referer"         = $signInUrl
    }
    $credentialJson = [PSCustomObject]@{
        username     = $Username
        password     = (ConvertTo-GCPlainText -SecureString $Password)
        rememberMe   = $true
        captchaToken = ""
    } | ConvertTo-Json -Compress
    # Send bytes: Windows PowerShell 5.1 does not encode string bodies as UTF-8.
    $loginResponse = Invoke-GCHttp -Uri ("{0}{1}/login?{2}" -f $script:GC.SsoBase, $Flow.ApiPath, $apiQuery) -Method POST `
        -WebSession $session -UserAgent $Flow.UserAgent -Headers $postHeaders -ContentType "application/json" `
        -Body ([System.Text.Encoding]::UTF8.GetBytes($credentialJson))
    $credentialJson = $null
    $result = Resolve-GCLoginResponse -Response $loginResponse

    if ($result.Status -eq "MfaRequired") {
        if ($null -eq $MfaCodeProvider) { throw "Garmin requires an MFA code, but no way to enter it is available." }
        $code = & $MfaCodeProvider $result.MfaMethod
        $mfaJson = [PSCustomObject]@{
            mfaMethod           = $result.MfaMethod
            mfaVerificationCode = ([string]$code).Trim()
            rememberMyBrowser   = $true
            reconsentList       = @()
            mfaSetup            = $false
        } | ConvertTo-Json -Compress
        $mfaResponse = Invoke-GCHttp -Uri ("{0}{1}/mfa/verifyCode?{2}" -f $script:GC.SsoBase, $Flow.ApiPath, $apiQuery) -Method POST `
            -WebSession $session -UserAgent $Flow.UserAgent -Headers $postHeaders -ContentType "application/json" `
            -Body ([System.Text.Encoding]::UTF8.GetBytes($mfaJson))
        $result = Resolve-GCLoginResponse -Response $mfaResponse
        if ($result.Status -ne "Success") { throw ("MFA verification failed ({0} {1})." -f $result.Status, $result.Message) }
    }

    switch ($result.Status) {
        "Success" {
            if ([string]::IsNullOrEmpty($result.Ticket)) { throw "RETRYABLE: login succeeded but Garmin returned no service ticket" }
            return $result.Ticket
        }
        "InvalidCredentials" { throw "Invalid Garmin Connect username or password." }
        default { throw ("RETRYABLE: {0} login: {1} {2}" -f $Flow.Name, $result.Status, $result.Message) }
    }
}

function Connect-GarminConnect {
    <#
    Returns a token object (AccessToken, RefreshToken, ClientId). Uses the token cache when possible and
    falls back to a full login (all login flows in order). The password provider is only invoked when a full
    login is needed, so a valid token cache means no password prompt.
    #>
    param(
        [Parameter(Mandatory)][string]$Username,
        [Parameter(Mandatory)][scriptblock]$PasswordProvider,
        [string]$TokenCachePath,
        [scriptblock]$MfaCodeProvider
    )
    if ($TokenCachePath) {
        $cached = Read-GCTokenCache -Path $TokenCachePath -Username $Username
        if ($cached) {
            $expiry = Get-GCTokenExpiry -Token $cached.AccessToken
            if ($expiry -and $expiry -gt [DateTime]::UtcNow.AddMinutes(5)) {
                Write-Host "INFO - Using cached Garmin Connect access token (valid until $($expiry.ToLocalTime()))."
                return $cached
            }
            $refreshed = Update-GCDiToken -Token $cached
            if ($refreshed) {
                Write-Host "INFO - Refreshed cached Garmin Connect access token."
                Save-GCTokenCache -Path $TokenCachePath -Token $refreshed -Username $Username
                return $refreshed
            }
            Write-Host "INFO - Cached token could not be refreshed, performing a full login."
        }
    }

    $Password = & $PasswordProvider
    if ($null -eq $Password -or $Password.Length -eq 0) { throw "No Garmin Connect password specified." }
    $errors = @()
    foreach ($flow in $script:GC.LoginFlows) {
        try {
            Write-Verbose "Trying Garmin SSO login flow '$($flow.Name)'"
            $ticket = Invoke-GCSsoLogin -Flow $flow -Username $Username -Password $Password -MfaCodeProvider $MfaCodeProvider
            $token = Get-GCDiToken -ServiceTicket $ticket -ServiceUrl $flow.ServiceUrl
            if ($TokenCachePath) { Save-GCTokenCache -Path $TokenCachePath -Token $token -Username $Username }
            return $token
        }
        catch {
            $message = $_.Exception.Message
            if ($message -notlike "RETRYABLE:*") { throw }
            Write-Verbose $message
            $errors += $message.Substring(11)
        }
    }
    throw ("All Garmin login flows failed. Garmin is probably rate limiting or blocking this client/network, please try again later. Details: {0}" -f ($errors -join "; "))
}
#endregion

#region Garmin Connect API
function Invoke-GCApi {
    <#
    Calls connectapi.garmin.com with the bearer token. Retries on HTTP 429 and 5xx with exponential back-off and
    refreshes the token once on HTTP 401. The token object is updated in place after a refresh.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$Token,
        [string]$Accept = "application/json",
        [scriptblock]$OnTokenRefreshed
    )
    $refreshed = $false
    for ($attempt = 0; $attempt -le $script:GC.MaxRetries; $attempt++) {
        $headers = $script:GC.NativeHeaders.Clone()
        $headers["Authorization"] = "Bearer $($Token.AccessToken)"
        $headers["Accept"] = $Accept
        $response = Invoke-GCHttp -Uri ($script:GC.ConnectApiBase + $Path) -Headers $headers -UserAgent $script:GC.NativeUserAgent

        if ($response.StatusCode -ge 200 -and $response.StatusCode -lt 300) { return $response }

        if ($response.StatusCode -eq 401 -and -not $refreshed) {
            $refreshed = $true
            $newToken = Update-GCDiToken -Token $Token
            if ($null -eq $newToken) { throw "Garmin rejected the access token (HTTP 401) and it could not be refreshed. Run the script again to log in." }
            $Token.AccessToken = $newToken.AccessToken
            $Token.RefreshToken = $newToken.RefreshToken
            $Token.ClientId = $newToken.ClientId
            if ($OnTokenRefreshed) { & $OnTokenRefreshed $Token }
            $attempt--
            continue
        }

        if (($response.StatusCode -eq 429 -or $response.StatusCode -ge 500) -and $attempt -lt $script:GC.MaxRetries) {
            $wait = [Math]::Pow(2, $attempt + 1) * 5
            $retryAfter = $response.Headers["Retry-After"]
            $retryAfterSeconds = 0
            if ($retryAfter -and [int]::TryParse($retryAfter, [ref]$retryAfterSeconds)) { $wait = [Math]::Min([Math]::Max($retryAfterSeconds, 1), 300) }
            Write-Warning ("WARNING - Garmin responded with HTTP {0} on {1}, retrying in {2} seconds..." -f $response.StatusCode, $Path, $wait)
            Start-Sleep -Seconds $wait
            continue
        }

        $hint = ""
        if ($response.StatusCode -eq 429) { $hint = " Your requests are rate limited, please try again later." }
        if ($response.StatusCode -eq 403) { $hint = " Garmin may be blocking this client or network." }
        throw ("Garmin Connect API call {0} failed with HTTP {1}.{2}" -f $Path, $response.StatusCode, $hint)
    }
}

function Get-GCActivityList {
    <#
    Returns all activities with an activity id greater than NewerThanActivityId (0 = all).
    The list is sorted by start time, not by id: an old activity that is uploaded later gets a new (higher) id but
    appears further down the list. Therefore all pages are always read and filtered on id, as in version 1.x.
    #>
    param(
        [Parameter(Mandatory)]$Token,
        [long]$NewerThanActivityId = 0,
        [scriptblock]$OnTokenRefreshed
    )
    $activities = New-Object System.Collections.Generic.List[object]
    $start = 0
    while ($true) {
        $path = "{0}?limit={1}&start={2}" -f $script:GC.ActivitySearchPath, $script:GC.PageSize, $start
        Write-Verbose "Activity list: $path"
        $response = Invoke-GCApi -Path $path -Token $Token -OnTokenRefreshed $OnTokenRefreshed
        $json = ConvertFrom-GCJsonSafe -Content $response.Content
        if ($null -eq $json -and -not [string]::IsNullOrWhiteSpace($response.Content) -and $response.Content.Trim() -ne "[]") {
            throw "Unexpected (non-JSON) answer from the Garmin activity list."
        }
        $page = @($json | Where-Object { $null -ne $_ })
        foreach ($activity in $page) {
            if ([long]$activity.activityId -gt $NewerThanActivityId) { $activities.Add($activity) }
        }
        if ($page.Count -lt $script:GC.PageSize) { break }
        $start += $script:GC.PageSize
    }
    return , $activities.ToArray()
}

function Save-GCActivityFile {
    <#
    Downloads one activity and stores it in the destination. FIT files are delivered by Garmin as a zip that
    can contain more than one file. Returns the number of files written.
    #>
    param(
        [Parameter(Mandatory)]$Activity,
        [Parameter(Mandatory)]$Token,
        [Parameter(Mandatory)][ValidateSet("FIT", "TCX", "GPX", "KML")][string]$FileType,
        [Parameter(Mandatory)][string]$Destination,
        [Parameter(Mandatory)][string]$TempDir,
        [bool]$Overwrite,
        [scriptblock]$OnTokenRefreshed
    )
    $extension = $FileType.ToLowerInvariant()
    $baseName = Get-GCSafeFileName -ActivityId $Activity.activityId -StartTimeLocal ([datetime]$Activity.startTimeLocal) `
        -ActivityType $Activity.activityType.typeKey -ActivityName $Activity.activityName

    $path = $script:GC.DownloadPaths[$FileType] + $Activity.activityId
    $response = Invoke-GCApi -Path $path -Token $Token -Accept "*/*" -OnTokenRefreshed $OnTokenRefreshed
    if ($null -eq $response.Bytes -or $response.Bytes.Length -eq 0) { throw "Garmin returned an empty file for activity $($Activity.activityId)." }

    $files = @()
    if ($FileType -eq "FIT") {
        $zipPath = Join-Path -Path $TempDir -ChildPath "$($Activity.activityId).zip"
        [System.IO.File]::WriteAllBytes($zipPath, $response.Bytes)
        Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
        try {
            $entries = @($zip.Entries | Where-Object { $_.Length -gt 0 -and -not $_.FullName.EndsWith("/") })
            $count = 0
            foreach ($entry in $entries) {
                $count++
                $suffix = ""
                if ($entries.Count -gt 1) { $suffix = "-$count" }
                $tempFile = Join-Path -Path $TempDir -ChildPath ("{0}{1}.{2}" -f $baseName, $suffix, $extension)
                # Only the entry name is used, never the path inside the zip (prevents path traversal).
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $tempFile, $true)
                $files += $tempFile
            }
        }
        finally { $zip.Dispose(); Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue }
    }
    else {
        $tempFile = Join-Path -Path $TempDir -ChildPath ("{0}.{1}" -f $baseName, $extension)
        [System.IO.File]::WriteAllBytes($tempFile, $response.Bytes)
        $files += $tempFile
    }

    $written = 0
    foreach ($file in $files) {
        $fileName = Split-Path -Path $file -Leaf
        $target = Join-Path -Path $Destination -ChildPath $fileName
        if ((Test-Path -LiteralPath $target) -and -not $Overwrite) {
            Write-Warning "WARNING - Skipping file $fileName because it already exists in the destination directory."
            Remove-Item -LiteralPath $file -Force
            continue
        }
        Write-Host "INFO - Writing file $target"
        Move-Item -LiteralPath $file -Destination $target -Force
        $written++
    }
    return $written
}
#endregion

#region Settings
function Read-GCUserSettings {
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path) -or -not (Test-Path -LiteralPath $Path)) { return $null }
    try { return ([xml](Get-Content -LiteralPath $Path -Raw)).GCUserSettings }
    catch { throw "The user settings file $Path is not valid XML: $($_.Exception.Message)" }
}

function Resolve-GCSetting {
    <# Script parameter wins, then the XML value, then the default. Values are validated against the allowed set. #>
    param([string]$Name, [string]$ParameterValue, [string]$XmlValue, [string]$Default, [string[]]$Allowed)
    $value = $Default; $source = "default"
    if (-not [string]::IsNullOrWhiteSpace($XmlValue)) { $value = $XmlValue.Trim(); $source = "user settings XML file" }
    if (-not [string]::IsNullOrWhiteSpace($ParameterValue)) { $value = $ParameterValue; $source = "script parameters" }
    if ($Allowed) {
        $match = $Allowed | Where-Object { $_ -eq $value } | Select-Object -First 1
        if (-not $match) { throw ("Invalid value '{0}' for {1} (from {2}). Allowed values: {3}" -f $value, $Name, $source, ($Allowed -join ", ")) }
        $value = $match
    }
    Write-Verbose "$Name = $value (from $source)"
    return $value
}

function Expand-GCEnvironmentPath {
    <# Supports '$env:NAME\folder' in the XML destination (kept for compatibility) and %NAME% / $HOME style values. #>
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path)) { return $Path }
    $expanded = [regex]::Replace($Path, '\$env:([A-Za-z_][A-Za-z0-9_]*)', {
            param($m) [Environment]::GetEnvironmentVariable($m.Groups[1].Value)
        })
    return [Environment]::ExpandEnvironmentVariables($expanded)
}
#endregion

#region Main
function Invoke-GCActivityExport {
    param([hashtable]$Parameters)

    $settingsPath = $Parameters.UserSettingsFile
    if ([string]::IsNullOrEmpty($settingsPath)) { $settingsPath = Join-Path -Path $PSScriptRoot -ChildPath "GCUserSettings.xml" }
    $userSettings = Read-GCUserSettings -Path $settingsPath
    if ($userSettings) { Write-Host "INFO - Using user settings file $settingsPath" }
    else { Write-Verbose "No user settings file found at $settingsPath" }

    $xml = @{ Username = $null; Password = $null; FileType = $null; DownloadOption = $null; Destination = $null; Overwrite = $null }
    if ($userSettings) {
        $xml.Username = $userSettings.Credentials.UserName
        $xml.Password = $userSettings.Credentials.Password
        $xml.FileType = $userSettings.ActivityFileType
        $xml.DownloadOption = $userSettings.DownloadOption
        $xml.Destination = $userSettings.Folders.Destination.Folder
        $xml.Overwrite = $userSettings.Folders.Destination.Overwrite
    }

    $fileType = Resolve-GCSetting -Name "ActivityFileType" -ParameterValue $Parameters.ActivityFileType -XmlValue $xml.FileType -Default "FIT" -Allowed "FIT", "TCX", "GPX", "KML"
    $downloadOption = Resolve-GCSetting -Name "DownloadOption" -ParameterValue $Parameters.DownloadOption -XmlValue $xml.DownloadOption -Default "New" -Allowed "New", "All"
    $overwrite = Resolve-GCSetting -Name "Overwrite" -ParameterValue $Parameters.Overwrite -XmlValue $xml.Overwrite -Default "No" -Allowed "Yes", "No"
    $destination = Resolve-GCSetting -Name "Destination" -ParameterValue $Parameters.Destination -XmlValue $xml.Destination -Default ""
    $destination = Expand-GCEnvironmentPath -Path $destination
    $username = Resolve-GCSetting -Name "Username" -ParameterValue $Parameters.Username -XmlValue $xml.Username -Default ""

    if ([string]::IsNullOrEmpty($destination) -or -not (Test-Path -LiteralPath $destination -PathType Container)) {
        if (-not [string]::IsNullOrEmpty($destination)) { Write-Warning "WARNING - Destination '$destination' does not exist." }
        if ((Test-GCIsWindows) -and [Environment]::UserInteractive) {
            Add-Type -AssemblyName System.Windows.Forms
            $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
            $folderBrowser.Description = "Select the destination folder for the Garmin Connect activities"
            if ($folderBrowser.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { throw "No destination folder selected." }
            $destination = $folderBrowser.SelectedPath
        }
        else { throw "Specify an existing destination folder with -Destination or in the user settings XML file." }
    }
    $destination = (Resolve-Path -LiteralPath $destination).ProviderPath

    if ([string]::IsNullOrEmpty($username)) { $username = Read-Host -Prompt "Enter your Garmin Connect username (e-mail address)" }
    if (-not [string]::IsNullOrEmpty($Parameters.Password)) { $securePassword = ConvertTo-SecureString -String $Parameters.Password -AsPlainText -Force }
    elseif (-not [string]::IsNullOrEmpty($xml.Password)) { $securePassword = ConvertTo-SecureString -String $xml.Password -AsPlainText -Force }
    else { $securePassword = $null }

    $tokenCachePath = $null
    if (-not $Parameters.NoTokenCache) {
        $tokenCachePath = $Parameters.TokenCacheFile
        if ([string]::IsNullOrEmpty($tokenCachePath)) { $tokenCachePath = Get-GCDefaultTokenCachePath }
    }

    Write-Host "INFO - Starting processing $fileType files from Garmin Connect with the following parameters:"
    Write-Host "- Activity File Type = $fileType"
    Write-Host "- Download Option = $downloadOption"
    Write-Host "- Destination = $destination"
    Write-Host "- Username = $username"
    Write-Host "- Overwrite = $overwrite"

    # Delta cookie
    $cookieFile = Join-Path -Path $destination -ChildPath ".GCDownloadStatus$fileType.cookie"
    $lastActivityId = [long]0
    if ($downloadOption -eq "New") {
        $deltaValue = $null
        if (Test-Path -LiteralPath $cookieFile) { $deltaValue = (Get-Content -LiteralPath $cookieFile -Raw).Trim() }
        if ($deltaValue -match '^[0-9]+$') { $lastActivityId = [long]$deltaValue }
        else {
            $question = "A valid delta cookie was not found in $destination, continue using a full export?"
            Write-Warning "WARNING - $question"
            $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))
            if ($Host.UI.PromptForChoice("Garmin Connect Activity Export", $question, $choices, 0) -ne 0) {
                Write-Host "INFO - Stopped by user."
                return
            }
            $downloadOption = "All"
        }
    }

    # Authenticate
    Write-Host "INFO - Connecting to Garmin Connect for user $username" -ForegroundColor Gray
    $mfaProvider = { param($method) Read-Host -Prompt "Enter the Garmin Connect MFA code (sent by $method)" }
    if ($null -ne $securePassword) { $passwordProvider = { $securePassword }.GetNewClosure() }
    else { $passwordProvider = { Read-Host -Prompt "Enter the Garmin Connect password for $username" -AsSecureString }.GetNewClosure() }
    $token = Connect-GarminConnect -Username $username -PasswordProvider $passwordProvider -TokenCachePath $tokenCachePath -MfaCodeProvider $mfaProvider
    $onRefresh = $null
    if ($tokenCachePath) { $onRefresh = { param($t) Save-GCTokenCache -Path $tokenCachePath -Token $t -Username $username }.GetNewClosure() }

    # Activity list
    Write-Host "INFO - Retrieving current status of your activities in Garmin Connect, please wait..."
    $activities = Get-GCActivityList -Token $token -NewerThanActivityId $lastActivityId -OnTokenRefreshed $onRefresh
    Write-Host "SUCCESS - Successfully connected to Garmin Connect" -ForegroundColor Green
    if ($activities.Count -eq 0) {
        Write-Host "INFO - No new activities found."
        return
    }
    Write-Host ("INFO - {0} activities to process, please wait..." -f $activities.Count)

    # Download, oldest first so an interrupted run can be resumed.
    $tempDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "GarminConnectActivityExportTMP"
    if (-not (Test-Path -LiteralPath $tempDir)) { $null = New-Item -Path $tempDir -ItemType Directory -Force }
    $sorted = $activities | Sort-Object -Property { [long]$_.activityId }
    $exported = 0; $failed = 0; $newestContiguous = $lastActivityId
    foreach ($activity in $sorted) {
        try {
            $null = Save-GCActivityFile -Activity $activity -Token $token -FileType $fileType -Destination $destination `
                -TempDir $tempDir -Overwrite ($overwrite -eq "Yes") -OnTokenRefreshed $onRefresh
            $exported++
            if ($failed -eq 0) { $newestContiguous = [long]$activity.activityId }
        }
        catch {
            $failed++
            Write-Error ("ERROR - Activity {0} could not be exported: {1}" -f $activity.activityId, $_.Exception.Message) -ErrorAction Continue
            if ($_.Exception.Message -like "*rate limited*") {
                Write-Warning "WARNING - Stopping because Garmin is rate limiting. Run the script again later to continue."
                break
            }
        }
    }

    # Delta cookie: only advance up to the last activity before the first failure, so failed activities are retried next run.
    if ($newestContiguous -gt $lastActivityId) {
        try {
            if (Test-Path -LiteralPath $cookieFile) { (Get-Item -LiteralPath $cookieFile -Force).Attributes = "Normal" }
            Set-Content -LiteralPath $cookieFile -Value $newestContiguous -Force
            if (Test-GCIsWindows) { (Get-Item -LiteralPath $cookieFile -Force).Attributes = "Hidden" }
        }
        catch { Write-Error "ERROR - Unable to write the delta file for later use: $($_.Exception.Message)" -ErrorAction Continue }
    }
    if ($failed -gt 0) { Write-Warning ("WARNING - Exported {0} activities to {1}, {2} failed." -f $exported, $destination, $failed) }
    else { Write-Host ("INFO - Finished exporting {0} activities from Garmin Connect to {1}. Delta file successfully stored." -f $exported, $destination) }
}

# Only run when executed, not when dot-sourced (for example by the Pester tests).
if ($MyInvocation.InvocationName -ne ".") {
    if ($PSVersionTable.PSEdition -eq "Desktop") {
        # Windows PowerShell 5.1 does not enable TLS 1.2 by default on every system.
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    }
    try {
        Invoke-GCActivityExport -Parameters $PSBoundParameters
    }
    catch {
        Write-Error ("ERROR - {0}" -f $_.Exception.Message) -ErrorAction Continue
        Write-Verbose ($_ | Out-String)
        exit 1
    }
}
#endregion
