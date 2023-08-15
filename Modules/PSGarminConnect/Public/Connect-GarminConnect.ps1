function Connect-GarminConnect {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        [System.String]$Username,
        [parameter(Mandatory = $true)]
        [System.Security.SecureString]$Password,
        [parameter(Mandatory = $false)]
        [switch]$Force
    )

    begin {
        $Config = GetPsGarminConnectConfig
    }
    process {
        if (-not [string]::IsNullOrEmpty($($Config.WebSession)) -and -not $Force){
                Write-Host "INFO - Already connected. To reconnect, use the -Force parameter" -ForegroundColor Gray
        }
        else {
            #Authenticate
            Write-Host "INFO - Connecting to Garmin Connect for user $Username" -ForegroundColor Gray
            $BaseLogin = Invoke-WebRequest -Uri $($Config.BaseURLs.BaseLoginURL) -WebSession $($Config.WebSession)
            $LoginForm = $BaseLogin.Forms[0]
            $LoginForm.Fields["username"] = "$Username"
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
            $LoginForm.Fields["password"] = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
            $Header = @{
                "origin"                    = "https://sso.garmin.com";
                "authority"                 = "connect.garmin.com"
                "scheme"                    = "https"
                "path"                      = "/signin/"
                "pragma"                    = "no-cache"
                "cache-control"             = "no-cache"
                "dnt"                       = "1"
                "upgrade-insecure-requests" = "1"
                "user-agent"                = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.182 Safari/537.36 Edg/88.0.705.81"
                "accept"                    = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
                "sec-fetch-site"            = "cross-site"
                "sec-fetch-mode"            = "navigate"
                "sec-fetch-user"            = "?1"
                "sec-fetch-dest"            = "document"
                "accept-language"           = "en,en-US;q=0.9,nl;q=0.8"
            }
            $Service = "service=https%3A%2F%2Fconnect.garmin.com%2Fmodern%2F"
            $BaseLogin = Invoke-RestMethod -Uri ($($Config.BaseURLs.BaseLoginURL) + "?" + $Service) -WebSession $($Config.WebSession) -Method POST -Body $LoginForm.Fields -Headers $Header
            Clear-Variable LoginForm -Force

            #Get Cookies
            $Cookies = $GarminConnectSession.Cookies.GetCookies($($Config.BaseURLs.BaseLoginURL))

            <#Show Cookies
            foreach ($cookie in $Cookies) {
                # You can get cookie specifics, or just use $cookie
                # This gets each cookie's name and value
                Write-Host "$($cookie.name) = $($cookie.value)"
            }#>

            #Get SSO cookie
            $SSOCookie = $Cookies | Where-Object name -EQ "CASTGC" | Select-Object value -ExpandProperty value
            if ($SSOCookie.Length -lt 1) {
                Write-Error "ERROR - No valid SSO cookie found, wrong credentials?" -ErrorAction Stop
            }

            #Authenticate by using cookie
            $PostLogin = Invoke-RestMethod -Uri ($($Config.PostLoginURL) + "?ticket=" + $SSOCookie) -WebSession $($Config.WebSession)
            $PostLogin

        }
    }
    end {}
}