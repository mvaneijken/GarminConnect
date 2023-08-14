#Get public and private function definition files.
$Public = @(Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue)
$Private = @(Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue)

#Dot source the files
Foreach ($import in @($Public + $Private)) {
    TRY {
        . $import.fullname
    }
    CATCH {
        Write-Error -Message  "Failed to import function $($import.fullname): $_"
    }
}

# Export all the functions
Export-ModuleMember -Function $Public.Basename -Alias *

$SessionVariableName = Get-GarminSessionVariable
New-Variable 


$ModuleSettings = [PSCustomObject]@{
    BaseURLs = [PSCustomObject]@{
        BaseLoginURL       = "https://sso.garmin.com/sso/login?service=https%3A%2F%2Fconnect.garmin.com%2Fmodern%2F&webhost=olaxpw-conctmodern004&source=https%3A%2F%2Fconnect.garmin.com%2Fnl-NL%2Fsignin&redirectAfterAccountLoginUrl=https%3A%2F%2Fconnect.garmin.com%2Fmodern%2F&redirectAfterAccountCreationUrl=https%3A%2F%2Fconnect.garmin.com%2Fmodern%2F&gauthHost=https%3A%2F%2Fsso.garmin.com%2Fsso&locale=nl_NL&id=gauth-widget&cssUrl=https%3A%2F%2Fstatic.garmincdn.com%2Fcom.garmin.connect%2Fui%2Fcss%2Fgauth-custom-v1.2-min.css&clientId=GarminConnect&rememberMeShown=true&rememberMeChecked=false&createAccountShown=true&openCreateAccount=false&usernameShown=false&displayNameShown=false&consumeServiceTicket=false&initialFocus=true&embedWidget=false&generateExtraServiceTicket=false&globalOptInShown=false&globalOptInChecked=false"
        PostLoginURL       = "https://connect.garmin.com/modern/"
        ActivitySearchURL  = "https://connect.garmin.com/proxy/activity-search-service-1.2/json/activities?"
        GPXActivityBaseURL = "https://connect.garmin.com/proxy/activity-service-1.1/gpx/activity/"
        TCXActivityBaseURL = "https://connect.garmin.com/proxy/activity-service-1.1/tcx/activity/"
        FITActivityBaseURL = "https://connect.garmin.com/proxy/download-service/files/activity/"
    }
}
