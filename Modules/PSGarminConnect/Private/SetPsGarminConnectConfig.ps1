function SetPsGarminConnectConfig {
    "Module Settings variable: Global:{0}" -f (GetPsGarminConnectConfig -Name) | Write-Verbose
    if ((Get-Variable | Where-Object { $_.Name -eq (GetPsGarminConnectConfig -Name) } | Measure-Object).Count -eq 0) {

        $ModuleConfig = [PSCustomObject]@{
            BaseURLs   = [PSCustomObject]@{
                BaseLoginURL       = "https://sso.garmin.com/sso/login?service=https%3A%2F%2Fconnect.garmin.com%2Fmodern%2F&webhost=olaxpw-conctmodern004&source=https%3A%2F%2Fconnect.garmin.com%2Fnl-NL%2Fsignin&redirectAfterAccountLoginUrl=https%3A%2F%2Fconnect.garmin.com%2Fmodern%2F&redirectAfterAccountCreationUrl=https%3A%2F%2Fconnect.garmin.com%2Fmodern%2F&gauthHost=https%3A%2F%2Fsso.garmin.com%2Fsso&locale=nl_NL&id=gauth-widget&cssUrl=https%3A%2F%2Fstatic.garmincdn.com%2Fcom.garmin.connect%2Fui%2Fcss%2Fgauth-custom-v1.2-min.css&clientId=GarminConnect&rememberMeShown=true&rememberMeChecked=false&createAccountShown=true&openCreateAccount=false&usernameShown=false&displayNameShown=false&consumeServiceTicket=false&initialFocus=true&embedWidget=false&generateExtraServiceTicket=false&globalOptInShown=false&globalOptInChecked=false"
                PostLoginURL       = "https://connect.garmin.com/modern/"
                ActivitySearchURL  = "https://connect.garmin.com/proxy/activity-search-service-1.2/json/activities?"
                GPXActivityBaseURL = "https://connect.garmin.com/proxy/activity-service-1.1/gpx/activity/"
                TCXActivityBaseURL = "https://connect.garmin.com/proxy/activity-service-1.1/tcx/activity/"
                FITActivityBaseURL = "https://connect.garmin.com/proxy/download-service/files/activity/"
            }
            WebSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        }
        "Module Settings: {0}" -f ($ModuleConfig | ConvertTo-Json ) | Write-Verbose
        New-Variable -Name (GetPsGarminConnectConfig -Name) -Force -Scope Global -Value $ModuleConfig
    }
    else {
        $ModuleConfig = Get-Variable -Name (GetPsGarminConnectConfig -Name) -Scope Global -ValueOnly
        if ($null -eq $ModuleConfig.WebSession) {
            $ModuleConfig.WebSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
            Set-Variable -Name (GetPsGarminConnectConfig -Name) -Force -Scope Global -Value $ModuleConfig
        }
    }
}