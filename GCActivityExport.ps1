# Garmin Connect Activity Export
# Mark van Eijken 2022
# Version 1.3.2
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
# The scripts does the following:
# - Downloads activity files from garmin in FIT, TCX or GPX format.
# - Supports delta download

PARAM(
    [CmdletBinding()]
    [ValidateSet("FIT", "TCX", "GPX")]
    $ActivityFileType = "FIT",
    [ValidateSet("All", "New")]
    $DownloadOption = "New",
    [parameter(Mandatory = $false)]
    $Destination,
    [parameter(Mandatory = $false)]
    $Username,
    [parameter(Mandatory = $false)]
    $Password,
    [ValidateSet("Yes", "No")]
    $Overwrite = "No"<#,
    [ValidateRange(1,999999999999999999)]
    $MaxItems = ""#>
)

#General parameters and settings
$error.Clear()
$ErrorActionPreference = "Stop"

#Read settings XML
try {
    $RAWProgramSettingsXML = Get-Content ".\GCProgramSettings.xml"
    $RAWProgramSettingsXML = $RAWProgramSettingsXML.replace("&", "###")
    [xml]$ProgramSettingsXML = $RAWProgramSettingsXML

    $RAWUserSettingsXML = Get-Content ".\GCUserSettings.xml"
    $RAWUserSettingsXML = $RAWUserSettingsXML.replace("&", "###")
    [xml]$UserSettingsXML = $RAWUserSettingsXML
}
catch {
    Write-Error "ERROR - Error occured reading XML files:`n$error"
}

##Checks
if ([string]::IsNullOrEmpty($username) -or [string]::IsNullOrEmpty($Password)) {
    #Check for Garmin Connect Activity Export - User Settings XML file
    Write-Host "INFO - Using Garmin Connect credentials from Garmin Connect Activity Export - User Settings XML file"
    $Username = $UserSettingsXML.GCUserSettings.Credentials.UserName
    $password = $UserSettingsXML.GCUserSettings.Credentials.Password
}
else {
    Write-Host "INFO - Using Garmin Connect credentials from script parameters"
}

#DestinationCheck
if ([string]::IsNullOrEmpty($Destination)) {
    #Check for Garmin Connect Activity Export - User Settings XML file
    Write-Host "INFO - Using Garmin Connect Destination from Garmin Connect Activity Export - User Settings XML file"
    $Destination = $UserSettingsXML.GCUserSettings.Folders.Destination.Folder

    #A work-arround to support system environment variables from the XML file.
    if ($Destination -like '$env:*') {
        $PsEnv = $Destination.split("\")[0]
        $PsEnvVariable = $PsEnv.split(":")[1]
        $PsEnvPath = Get-ChildItem env: | Where-Object Name -EQ $PsEnvVariable | Select-Object value -ExpandProperty value
        $destination = $Destination.Replace($PsEnv, $PsEnvPath)
    }
}
else {
    Write-Host "INFO - Using Garmin Connect Destination from script parameters"
}
$invaliddestination = $false
if (!(Test-Path $Destination -ErrorAction SilentlyContinue)) {
    Write-Error "ERROR - Incorrect Destination configured in the user settings XML file. You will be asked to select a destination"
    $invaliddestination = $true
}
else {
    Write-Host "INFO - Destination directory $Destination correct."
}

#ActivityFileType Check
if ([string]::IsNullOrEmpty($ActivityFileType)) {
    #Check for Garmin Connect Activity Export - User Settings XML file

    $ActivityFileType = $UserSettingsXML.GCUserSettings.ActivityFileType
    $ActivityFileTypeValidateSet = @("FIT", "TCX", "GPX")
    $ActivityFileTypeCorrect = $false
    foreach ($i in $ActivityFileTypeValidateSet) {
        if ($i -eq $ActivityFileType) {

            $ActivityFileTypeCorrect = $true
        }
    }
    if ($ActivityFileTypeCorrect -ne $true) {
        Write-Error "Incorrect activity filetype configured in the user settings XML file. Please use one of the following types:"
        foreach ($a in $ActivityFileTypeValidateSet) {
            Write-Host $a -ForegroundColor Red
        }
    }
    else {
        Write-Host "INFO - Using Garmin Connect ActivityFileType from Garmin Connect Activity Export - User Settings XML file"
    }
}
else {
    Write-Host "INFO - Using Garmin Connect ActivityFileType from script parameters"
}

#Download option check
if ([string]::IsNullOrEmpty($DownloadOption)) {
    #Check for Garmin Connect Activity Export - User Settings XML file
    Write-Host "INFO - Using Garmin Connect DownloadOption from Garmin Connect Activity Export - User Settings XML file"
    $DownloadOption = $UserSettingsXML.GCUserSettings.DownloadOption
    $DownloadOptionValidateSet = @("New", "All")
    $DownloadOptionCorrect = $false
    foreach ($i in $DownloadOptionValidateSet) {
        if ($i -eq $DownloadOption) {
            $DownloadOptionCorrect = $true
        }
    }
    if ($DownloadOptionCorrect -ne $true) {
        Write-Error "Incorrect download option configured in the user settings XML file. Please use one of the following types:"
        foreach ($d in $DownloadOptionValidateSet) {
            Write-Host $d -ForegroundColor Red
        }
    }
    else {
        Write-Host "INFO - Using Garmin Connect download option from Garmin Connect Activity Export - User Settings XML file"
    }
}
else {
    Write-Host "INFO - Using Garmin Connect DownloadOption from script parameters"
}

#Overwrite check
if ([string]::IsNullOrEmpty($Overwrite)) {
    #Check for Garmin Connect Activity Export - User Settings XML file
    Write-Host "INFO - Using Garmin Connect Overwrite from Garmin Connect Activity Export - User Settings XML file"
    $Overwrite = $UserSettingsXML.GCUserSettings.Folders.Destination.Overwrite
    $OverwriteValidateSet = @("Yes", "No")
    $OverwriteCorrect = $false
    foreach ($i in $OverwriteValidateSet) {
        if ($i -eq $Overwrite) {
            $OverwriteCorrect = $true
        }
    }
    if ($OverwriteCorrect -ne $true) {
        Write-Error "Incorrect Overwrite configured in the user settings XML file. Please use one of the following types:"
        foreach ($o in $OverwriteValidateSet) {
            Write-Host $o -ForegroundColor Red
        }
    }
    else {
        Write-Host "INFO - Using Garmin Connect Overwrite from Garmin Connect Activity Export - User Settings XML file"
    }
}
else {
    Write-Host "INFO - Using Garmin Connect Overwrite from script parameters"
}

#Get URLs neede from program settings xml
$ProgramSettingsXMLBaseURLNodes = $ProgramSettingsXML.GCProgramSettings.BaseURLs | Get-Member | Where-Object name -NotLike "#*" | Where-Object membertype -EQ "property" | Select-Object name -ExpandProperty name
foreach ($n in $ProgramSettingsXMLBaseURLNodes) {
    $variablename = $n
    New-Variable $n -Force
    if ([string]::IsNullOrEmpty($($ProgramSettingsXML.GCProgramSettings.BaseURLs.$n))) {
        Write-Error "Setting $n is empty in the Garmin Connect Activity Export - Program Settings XML file. Please correct"
    }
    else {
        Set-Variable -Name $variablename -Value  $($ProgramSettingsXML.GCProgramSettings.BaseURLs.$n).replace("###", "&")
        $n = $($ProgramSettingsXML.GCProgramSettings.BaseURLs.$n)
        #Get-Variable $variablename
    }
}

#Check if needed parameters are present
if ([string]::IsNullOrEmpty($Username)) { $Username = Read-Host -Prompt "Enter your username" }
if ([string]::IsNullOrEmpty($Password)) { $Password = Read-Host -Prompt "Enter the password for user $Username" }
if (([string]::IsNullOrEmpty($Destination) -or $invaliddestination -eq $true)) {
    Add-Type -AssemblyName System.Windows.Forms
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    [void]$FolderBrowser.ShowDialog()
    $Destination = $FolderBrowser.SelectedPath
}
else { if ($Destination.EndsWith("\")) { $Destination = $Destination.TrimEnd("\") } }

$CookieFilename = ".GCDownloadStatus$ActivityFileType.cookie"
$CookieFileFullPath = ($Destination + "\" + $CookieFilename)

#Write process information:
Write-Host "INFO - Starting processing $ActivityFileType files from Garmin Connect with the following parameters:"
Write-Host "- Activity File Type = $ActivityFileType"
Write-Host "- Download Option = $DownloadOption"
Write-Host "- Destination = $Destination"
Write-Host "- Username = $Username"
Write-Host "- Overwrite = $Overwrite"

#Authenticate
Write-Host "INFO - Connecting to Garmin Connect for user $Username" -ForegroundColor Gray
$BaseLogin = Invoke-WebRequest -Uri $BaseLoginURL -SessionVariable GarminConnectSession
$LoginForm = $BaseLogin.Forms[0]
$LoginForm.Fields["username"] = "$Username"
$LoginForm.Fields["password"] = "$Password"
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
$BaseLogin = Invoke-RestMethod -Uri ($BaseLoginURL + "?" + $Service) -WebSession $GarminConnectSession -Method POST -Body $LoginForm.Fields -Headers $Header

#Get Cookies
$Cookies = $GarminConnectSession.Cookies.GetCookies($BaseLoginURL)

<#Show Cookies
foreach ($cookie in $Cookies) {
     # You can get cookie specifics, or just use $cookie
     # This gets each cookie's name and value
     Write-Host "$($cookie.name) = $($cookie.value)"
}#>

#Get SSO cookie
$SSOCookie = $Cookies | Where-Object name -EQ "CASTGC" | Select-Object value -ExpandProperty value
if ($SSOCookie.Length -lt 1) {
    Write-Error "ERROR - No valid SSO cookie found, wrong credentials?"
    break
}

#Authenticate by using cookie
$PostLogin = Invoke-RestMethod -Uri ($PostLoginURL + "?ticket=" + $SSOCookie) -WebSession $GarminConnectSession

#Set the correct activity download URL for the selected type.
switch ($ActivityFileType) {
    'TCX' { $ActivityBaseURL = $TCXActivityBaseURL }
    'GPX' { $ActivityBaseURL = $GPXActivityBaseURL }
    Default { $ActivityBaseURL = $FITActivityBaseURL }
}

#Get activity pages and check if the connection is successfull
$ActivityList = @()
$PageSize = 100
$FirstRecord = 0
$Pages = 0
do {
    $SearchResults = Invoke-RestMethod -Uri $ActivitySearchURL"?limit=$PageSize&start=$FirstRecord" -Method get -WebSession $GarminConnectSession -ErrorAction SilentlyContinue -Headers @{
        "method"           = "GET"
        "authority"        = "connect.garmin.com"
        "scheme"           = "https"
        "accept"           = "application/json, text/javascript, */*; q=0.01"
        "dnt"              = "1"
        "x-requested-with" = "XMLHttpRequest"
        "user-agent"       = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.182 Safari/537.36 Edg/88.0.705.81"
        "nk"               = "NT"
        "sec-fetch-site"   = "same-origin"
        "sec-fetch-mode"   = "cors"
        "sec-fetch-dest"   = "empty"
        "referer"          = "https://connect.garmin.com/modern/activities"
        "accept-language"  = "en,en-US;q=0.9,nl;q=0.8"
    }
    $ActivityList += $SearchResults
    $FirstRecord = $FirstRecord + $PageSize
    $Pages++
}
until ($SearchResults.Count -eq 0)

if ($Pages -gt 0) { Write-Host "SUCCESS - Successfully connected to Garmin Connect" -ForegroundColor Green }
else {
    Write-Error "ERROR - Connection to Garmin Connect failed. Error:`n$($error[0])."
    break
}
#$TotalPages = $Pages

#Validate download option
if ($DownloadOption -eq "New") {
    $ErrorFound = $true
    if (Test-Path $CookieFileFullPath -ErrorAction SilentlyContinue) {
        $DeltaCookie = Get-Content $CookieFileFullPath
        if (($DeltaCookie -match '^[0-9]*$') -and (($DeltaCookie | Measure-Object -Line).lines -eq 1) -and (($DeltaCookie | Measure-Object -Word).words -eq 1)) {
            $ErrorFound = $false
        }
    }
    if ($ErrorFound -eq $true) {
        $msg = "A valid delta cookie not found, continue using a full import?"
        Write-Warning "WARNING - $msg"
        $question = $msg
        $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))
        $decision = $Host.UI.PromptForChoice($message, $question, $choices, 0)
        if ($decision -eq 0) { $DownloadOption = "All" }
        elseif ($decision -eq 1) {
            Write-Host "INFO - Stopped by user."
            break
        }
        else {
            Write-Host "INFO - Stopped by user."
            break
        }
    }
}

#Get activities
$Activities = @()
if ($DownloadOption -eq "New") {
    Write-Host "INFO - Retrieving current status of your activities in Garmin Connect, please wait..."
    foreach ($Activity in $ActivityList) {
        if ($($Activity.activityId) -gt $DeltaCookie) { $Activities += $Activity }
    }
}
else {
    Write-Host "INFO - Retrieving current status of your activities in Garmin Connect, please wait..."
    $Activities = $ActivityList
}

#Download activities in queue and unpack to destination location
Write-Host "INFO - Continue to process all retrieved activities, please wait..."

try {
    $TempDir = Join-Path -Path $env:temp -ChildPath GarminConnectActivityExportTMP
    $ActivityFileType = $ActivityFileType.tolower()
    if (!(Test-Path $TempDir)) { $null = New-Item -Path $TempDir -ItemType Directory -Force }
    $ActivityExportedCount = 0
    foreach ($Activity in $Activities) {
        #Download files
        $URL = $ActivityBaseURL + $($Activity.activityID) + "/"
        if ($ActivityFileType -eq "fit") {
            $OutputFileFullPath = Join-Path -Path $TempDir -ChildPath "$($Activity.activityID).zip"
        }
        else { $OutputFileFullPath = Join-Path -Path $TempDir -ChildPath "$($Activity.activityID).$ActivityFileType" }
        if (Test-Path $OutputFileFullPath) {
            #Allways overwrite temp files
            $null = Remove-Item $OutputFileFullPath -Force
        }
        Invoke-RestMethod -Uri $URL  -WebSession $GarminConnectSession -OutFile $OutputFileFullPath

        #Setting naming parameters for having the file to a more readable format
        $ActivityID = $Activity.activityId
        $ActivityName = $Activity.activityName
        $ActivityType = $Activity.activityType.typekey
        $ActivityBeginTimeStamp = (Get-Date $Activity.startTimeLocal).ToString("yyyy-MM-dd")
        $NamingMask = (("$ActivityID - $ActivityBeginTimeStamp - $ActivityType - $ActivityName").TrimEnd() -replace $_.name -replace '[^A-Za-z0-9-_\@\,\(\) \.\[\]]', '-')
        #Unzip the temporary files for FIT files and move all files to the destination location
        if ($ActivityFileType -eq "fit") {
            $Shell = New-Object -com shell.application
            $ZIP = $shell.NameSpace(“$OutputFileFullPath”)
            if ($zip.items().count -gt 1) {
                $Count = 0
                foreach ($Item in $ZIP.items()) {
                    $Count++
                    Write-Host "INFO - Downloading file $Destination$NamingMask-$Count.$ActivityFileType"
                    $Shell.Namespace(“$TempDir”).copyhere($Item, 0x14)
                    $DownloadedFileFullPath = Join-Path -Path $TempDir -ChildPath $($Item.name)
                    $FinalFileName = ($NamingMask + "-" + $Count + "." + $ActivityFileType)
                    $FinalFileNameTempFullPath = Join-Path -Path $TempDir -ChildPath $FinalFileName
                    if (Test-Path $FinalFileNameTempFullPath) {
                        #Allways overwrite temp files
                        $null = Remove-Item $FinalFileNameTempFullPath -Force
                    }
                    Rename-Item $DownloadedFileFullPath -NewName $FinalFileNameTempFullPath -Force
                    if ($Overwrite -eq "Yes") {
                        Move-Item $FinalFileNameTempFullPath -Destination $Destination -Force
                    }
                    else {
                        if (Test-Path (Join-Path -Path $Destination -ChildPath $FinalFileName)) {
                            Write-Warning "WARNING - Skipping file $FinalFileName because it allready exists in the destination directory."
                        }
                        else {
                            Move-Item $FinalFileNameTempFullPath -Destination $Destination
                        }
                    }

                }
            }

            elseif ($zip.items().count -eq 1) {
                Write-Host "INFO - Trying to create file $NamingMask.$ActivityFileType"
                foreach ($Item in $ZIP.items()) {
                    $Shell.Namespace(“$TempDir”).copyhere($zip.items(), 0x14)
                    $DownloadedFileFullPath = Join-Path -Path $TempDir -ChildPath $($Item.name)
                    $FinalFileName = ($NamingMask + "." + $ActivityFileType)
                    $FinalFileNameTempFullPath = Join-Path -Path $TempDir -ChildPath $FinalFileName
                    if (Test-Path $FinalFileNameTempFullPath) {
                        #Allways overwrite temp files
                        $null = Remove-Item $FinalFileNameTempFullPath -Force
                    }
                    Rename-Item $DownloadedFileFullPath -NewName $FinalFileNameTempFullPath -Force
                    if ($Overwrite -eq "Yes") {
                        Move-Item $FinalFileNameTempFullPath -Destination $Destination -Force
                    }
                    else {
                        if (Test-Path (Join-Path -Path $Destination -ChildPath $FinalFileName)) {
                            Write-Warning "WARNING - Skipping file $FinalFileName because it allready exists in the destination directory."
                        }
                        else {
                            Move-Item $FinalFileNameTempFullPath -Destination $Destination
                        }
                    }
                }
            }
        }
        #All other filetypes are considered as non-zip.
        else {
            Write-Host "INFO - Trying to create file $NamingMask.$ActivityFileType"
            $FinalFileName = ($NamingMask + "." + $ActivityFileType)
            $FinalFileNameTempFullPath = Join-Path -Path $TempDir -ChildPath $FinalFileName
            Rename-Item $OutputFileFullPath -NewName $FinalFileNameTempFullPath
            if ($Overwrite -eq "Yes") {
                Move-Item $FinalFileNameTempFullPath -Destination $Destination -Force
            }
            else {
                if (Test-Path (Join-Path -Path $Destination -ChildPath $FinalFileName)) {
                    Write-Warning "WARNING - Skipping file $FinalFileName because it allready exists in the destination directory."
                }
                else {
                    Move-Item $FinalFileNameTempFullPath -Destination $Destination
                }
            }
        }
        $ActivityExportedCount++
    }
}

catch {
    Write-Error "ERROR - An unexpected error occurred. See errordetails below:`n$($error[0])"
    break
}

#Finally write the hidden delta cookie for later use
if (($Activities[0].activityId).length -gt 0) {
    if (!(Test-Path $CookieFileFullPath)) { $null = New-Item $CookieFileFullPath -ItemType File -Force }
    $NewestActivity = ($Activities[0].activityId)
    try {

        $CookieFileFullPath = Join-Path -Path $Destination -ChildPath $CookieFilename
        (Get-Item $CookieFileFullPath -Force).Attributes = "Normal"
        $NewestActivity | Out-File $CookieFileFullPath -Force
        (Get-Item $CookieFileFullPath -Force).Attributes = "Hidden"
        Write-Host "INFO - Finished exporting $ActivityExportedCount activities from Garmin Connect. Delta file successfully stored."
    }
    catch {
        Write-Error "ERROR - Unable the write the delta file for later use, see error details below:`n$($error[0])"
    }
}
else { Write-Warning "No new activities found." }
