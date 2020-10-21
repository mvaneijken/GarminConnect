# Garmin Connect Activity Export
# Mark van Eijken 2018
# Version 1.2
# Version history:
# 1.0 - Initial version
# 1.1 - Fix: Garmin now expects parameters in the SSO url
#       Update: Added settings support in separate XML files
# 1.2 - Fix: Garmin removed the login action from the login response. It's now added as a static string.
#       Update: It seems that the json provided for an activity does not always contain all needed fields. Fallback is using only the basic information for creating a activity file.
# The scripts does the following:
# - Downloads activity files from garmin in FIT, TCX or GPX format.
# - Supports delta download

PARAM(
    [ValidateSet("FIT", "TCX", "GPX")]
    $ActivityFileType = "",
    [ValidateSet("All", "New")]
    $DownloadOption = "",
    [parameter(Mandatory = $false)]
    $Destination = "",
    [parameter(Mandatory = $false)]
    $Username = "",
    [parameter(Mandatory = $false)]
    $Password = "",
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
    $RAWProgramSettingsXML = get-content ".\Garmin Connect Activity Export - Program Settings.xml"
    $RAWProgramSettingsXML = $RAWProgramSettingsXML.replace("&", "###")
    [xml]$ProgramSettingsXML = $RAWProgramSettingsXML

    $RAWUserSettingsXML = get-content ".\Garmin Connect Activity Export - User Settings.xml"
    $RAWUserSettingsXML = $RAWUserSettingsXML.replace("&", "###")
    [xml]$UserSettingsXML = $RAWUserSettingsXML
}
catch {
    write-error "ERROR - Error occured reading XML files:`n$error"
}

##Checks
if ([string]::IsNullOrEmpty($username) -or [string]::IsNullOrEmpty($Password)) {
    #Check for Garmin Connect Activity Export - User Settings XML file
    write-host "INFO - Using Garmin Connect credentials from Garmin Connect Activity Export - User Settings XML file"
    $Username = $UserSettingsXML.GCUserSettings.Credentials.UserName
    $password = $UserSettingsXML.GCUserSettings.Credentials.Password
}
else {
    write-host "INFO - Using Garmin Connect credentials from script parameters"
}

#DestinationCheck
if ([string]::IsNullOrEmpty($Destination)) {
    #Check for Garmin Connect Activity Export - User Settings XML file
    write-host "INFO - Using Garmin Connect Destination from Garmin Connect Activity Export - User Settings XML file"
    $Destination = $UserSettingsXML.GCUserSettings.Folders.Destination.Folder

    #A work-arround to support system environment variables from the XML file.
    if ($Destination -like '$env:*') {
        $PsEnv = $Destination.split("\")[0]
        $PsEnvVariable = $PsEnv.split(":")[1]
        $PsEnvPath = Get-ChildItem env: | Where-Object Name -eq $PsEnvVariable | select value -ExpandProperty value
        $destination = $Destination.Replace($PsEnv, $PsEnvPath)
    }
}
else {
    write-host "INFO - Using Garmin Connect Destination from script parameters"
}
$invaliddestination = $false
if (!(test-path $Destination -ErrorAction SilentlyContinue)) {
    Write-error "ERROR - Incorrect Destination configured in the user settings XML file. You will be asked to select a destination"
    $invaliddestination = $true
}
else {
    write-host "INFO - Destination directory $Destination correct."
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
        write-error "Incorrect activity filetype configured in the user settings XML file. Please use one of the following types:"
        foreach ($a in $ActivityFileTypeValidateSet) {
            write-host $a -ForegroundColor Red
        }
    }
    else {
        write-host "INFO - Using Garmin Connect ActivityFileType from Garmin Connect Activity Export - User Settings XML file"
    }
}
else {
    write-host "INFO - Using Garmin Connect ActivityFileType from script parameters"
}

#Download option check
if ([string]::IsNullOrEmpty($DownloadOption)) {
    #Check for Garmin Connect Activity Export - User Settings XML file
    write-host "INFO - Using Garmin Connect DownloadOption from Garmin Connect Activity Export - User Settings XML file"
    $DownloadOption = $UserSettingsXML.GCUserSettings.DownloadOption
    $DownloadOptionValidateSet = @("New", "All")
    $DownloadOptionCorrect = $false
    foreach ($i in $DownloadOptionValidateSet) {
        if ($i -eq $DownloadOption) {
            $DownloadOptionCorrect = $true
        }
    }
    if ($DownloadOptionCorrect -ne $true) {
        write-error "Incorrect download option configured in the user settings XML file. Please use one of the following types:"
        foreach ($d in $DownloadOptionValidateSet) {
            write-host $d -ForegroundColor Red
        }
    }
    else {
        write-host "INFO - Using Garmin Connect download option from Garmin Connect Activity Export - User Settings XML file"
    }
}
else {
    write-host "INFO - Using Garmin Connect DownloadOption from script parameters"
}

#Overwrite check
if ([string]::IsNullOrEmpty($Overwrite)) {
    #Check for Garmin Connect Activity Export - User Settings XML file
    write-host "INFO - Using Garmin Connect Overwrite from Garmin Connect Activity Export - User Settings XML file"
    $Overwrite = $UserSettingsXML.GCUserSettings.Folders.Destination.Overwrite
    $OverwriteValidateSet = @("Yes", "No")
    $OverwriteCorrect = $false
    foreach ($i in $OverwriteValidateSet) {
        if ($i -eq $Overwrite) {
            $OverwriteCorrect = $true
        }
    }
    if ($OverwriteCorrect -ne $true) {
        write-error "Incorrect Overwrite configured in the user settings XML file. Please use one of the following types:"
        foreach ($o in $OverwriteValidateSet) {
            write-host $o -ForegroundColor Red
        }
    }
    else {
        write-host "INFO - Using Garmin Connect Overwrite from Garmin Connect Activity Export - User Settings XML file"
    }
}
else {
    write-host "INFO - Using Garmin Connect Overwrite from script parameters"
}

#Get URLs neede from program settings xml
$ProgramSettingsXMLBaseURLNodes = $ProgramSettingsXML.GCProgramSettings.BaseURLs | gm | where-object name -notlike "#*" | where-object membertype -eq "property" | select name -ExpandProperty name
foreach ($n in $ProgramSettingsXMLBaseURLNodes) {
    $variablename = $n
    New-Variable $n -Force
    if ([string]::IsNullOrEmpty($($ProgramSettingsXML.GCProgramSettings.BaseURLs.$n))) {
        write-error "Setting $n is empty in the Garmin Connect Activity Export - Program Settings XML file. Please correct"
    }
    else {
        Set-Variable -Name $variablename -value  $($ProgramSettingsXML.GCProgramSettings.BaseURLs.$n).replace("###", "&")
        $n = $($ProgramSettingsXML.GCProgramSettings.BaseURLs.$n)
        #Get-Variable $variablename
    }
}

#Check if needed parameters are present
if ([string]::IsNullOrEmpty($Username)) {$Username = Read-Host -Prompt "Enter your username"}
if ([string]::IsNullOrEmpty($Password)) {$Password = Read-Host -Prompt "Enter the password for user $Username"}
if (([string]::IsNullOrEmpty($Destination) -or $invaliddestination -eq $true)) {
    Add-Type -AssemblyName System.Windows.Forms
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    [void]$FolderBrowser.ShowDialog()
    $Destination = $FolderBrowser.SelectedPath
}
else {if ($Destination.EndsWith("\")) {$Destination = $Destination.TrimEnd("\")}}

$CookieFilename = ".GCDownloadStatus$ActivityFileType.cookie"
$CookieFileFullPath = ($Destination + "\" + $CookieFilename)

#Write process information:
write-host "INFO - Starting processing $ActivityFileType files from Garmin Connect with the following parameters:"
write-host "- Activity File Type = $ActivityFileType"
write-host "- Download Option = $DownloadOption"
write-host "- Destination = $Destination"
write-host "- Username = $Username"
write-host "- Overwrite = $Overwrite"

#Authenticate
write-host "INFO - Connecting to Garmin Connect for user $Username" -ForegroundColor Gray
$BaseLogin = Invoke-WebRequest -URI $BaseLoginURL -SessionVariable GarminConnectSession
$LoginForm = $BaseLogin.Forms[0]
$LoginForm.Fields["username"] = "$Username"
$LoginForm.Fields["password"] = "$Password"
$BaseLogin = Invoke-RestMethod -Uri ($BaseLoginURL + "</BaseLoginURL") -WebSession $GarminConnectSession -Method POST -Body $LoginForm.Fields

#Get Cookies
$Cookies = $GarminConnectSession.Cookies.GetCookies($BaseLoginURL)

<#Show Cookies
foreach ($cookie in $Cookies) {
     # You can get cookie specifics, or just use $cookie
     # This gets each cookie's name and value
     Write-Host "$($cookie.name) = $($cookie.value)"
}#>

#Get SSO cookie
$SSOCookie = $Cookies | Where-Object name -eq "CASTGC" | select value -ExpandProperty value
if ($SSOCookie.Length -lt 1) {
    write-error "ERROR - No valid SSO cookie found, wrong credentials?"
    break
}

#Authenticate by using cookie
$PostLogin = Invoke-RestMethod -Uri ($PostLoginURL + "?ticket=" + $SSOCookie) -WebSession $GarminConnectSession

#Set the correct activity download URL for the selected type.
switch ($ActivityFileType) {
    'TCX' {$ActivityBaseURL = $TCXActivityBaseURL}
    'GPX' {$ActivityBaseURL = $GPXActivityBaseURL}
    Default {$ActivityBaseURL = $FITActivityBaseURL}
}

#Get activity pages and check if the connection is successfull
$SearchResults = Invoke-RestMethod -Uri $ActivitySearchURL -method get -WebSession $GarminConnectSession -ErrorAction SilentlyContinue
if (!([string]::IsNullOrEmpty($($SearchResults.results.currentPage)))) {write-host "SUCCESS - Successfully connected to Garmin Connect" -ForegroundColor Green}
else {
    write-error "ERROR - Connection to Garmin Connect failed. Error:`n$($error[0])."
    break
}
$FirstActivityPage = $($SearchResults.results)
$TotalPages = $FirstActivityPage.totalPages

#Validate download option
if ($DownloadOption -eq "New") {
    $ErrorFound = $true
    if (test-path $CookieFileFullPath -ErrorAction SilentlyContinue) {
        $DeltaCookie = Get-Content $CookieFileFullPath
        if (($DeltaCookie -match '^[0-9]*$') -and (($DeltaCookie | Measure-Object -Line).lines -eq 1) -and (($DeltaCookie | Measure-Object -word).words -eq 1)) {
            $ErrorFound = $false
        }
    }
    if ($ErrorFound -eq $true) {
        $msg = "A valid delta cookie not found, continue using a full import?"
        write-warning "WARNING - $msg"
        $question = $msg
        $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList "&Yes"))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList "&No"))
        $decision = $Host.UI.PromptForChoice($message, $question, $choices, 0)
        if ($decision -eq 0) {$DownloadOption = "All"}
        elseif ($decision -eq 1) {
            write-host "INFO - Stopped by user."
            break
        }
        else {
            write-host "INFO - Stopped by user."
            break
        }
    }
}

#Get activities
if ($DownloadOption -eq "New") {
    write-host "INFO - Retrieving current status of your activities in Garmin Connect, please wait..."
    $Activities = @()
    $ActivityPages = @()
    $Count = 1
    do {
        $Query = "currentPage=$Count"
        $ActivityPages += (Invoke-RestMethod -Uri $ActivitySearchURL$Query -method get -WebSession $GarminConnectSession).results.activities
        $Count++
    }
    until (($Count -eq $Totalpages) -or ($($ActivityPages.activity.activityId) -contains $DeltaCookie))

    foreach ($Activity in $ActivityPages) {
        if ($($Activity.activity.activityId) -gt $DeltaCookie) {$Activities += $Activity}
    }
}
else {
    write-host "INFO - Retrieving current status of your activities in Garmin Connect, please wait..."
    $Activities = @()
    $Count = 1
    do {
        $Query = "currentPage=$count"
        $Activities += (Invoke-RestMethod -Uri $ActivitySearchURL$query -method get -WebSession $GarminConnectSession).results.activities
        $Count++
    }
    until ($Count -eq $TotalPages)
}

#Download activities in queu and unpack to destination location
write-host "INFO - Continue to process all retrieved activities, please wait..."

try {
    $TempDir = join-path -path $env:temp -childpath GarminConnectActivityExportTMP
    $ActivityFileType = $ActivityFileType.tolower()
    if (!(Test-Path $TempDir)) {$null = New-Item -Path $TempDir -ItemType Directory -Force}
    $ActivityExportedCount = 0
    foreach ($Activity in $Activities) {
        #Download files
        $URL = $ActivityBaseURL + $($activity.activity.activityID) + "/"
        if ($ActivityFileType -eq "fit") {
            $OutputFileFullPath = join-path -path $TempDir -ChildPath "$($activity.activity.activityID).zip"
        }
        else {$OutputFileFullPath = join-path -path $TempDir -ChildPath "$($activity.activity.activityID).$ActivityFileType"}
        if (Test-Path $OutputFileFullPath) {
            #Allways overwrite temp files
            $null = Remove-Item $OutputFileFullPath -Force
        }
        Invoke-RestMethod -Uri $URL  -WebSession $GarminConnectSession -OutFile $OutputFileFullPath

        #Setting naming parameters for having the file to a more readable format
        $ActivityID = $Activity.activity.activityId
        $ActivityName = $Activity.activity.activityName
        $ActivityType = ($Activity.activity.activityType.key).substring(0, 1).ToUpper() + ($Activity.activity.activityType.key).substring(1).tolower()

        $ActivitySummary = $Activity.activity.activitySummary
        if ($ActivitySummary) {
            $ActivityBeginTimeStamp = $Activity.activity.activitySummary.BeginTimestamp.value.Substring(0, 10)
        }
        else {
            $ActivityBeginTimeStamp = (get-date $Activity.activity.updatedDate.'$').ToString("yyyy-MM-dd")
        }
        $NamingMask = (("$ActivityID - $ActivityBeginTimeStamp - $ActivityType - $ActivityName").TrimEnd() -replace $_.name -replace '[^A-Za-z0-9-_\@\,\(\) \.\[\]]', '-')

        #Unzip the temporary files for FIT files and move all files to the destination location
        if ($ActivityFileType -eq "fit") {
            $Shell = new-object -com shell.application
            $ZIP = $shell.NameSpace(“$OutputFileFullPath”)
            if ($zip.items().count -gt 1) {
                $Count = 0
                foreach ($Item in $ZIP.items()) {
                    $Count++
                    write-host "INFO - Downloading file $Destination$NamingMask-$Count.$ActivityFileType"
                    $Shell.Namespace(“$TempDir”).copyhere($Item, 0x14)
                    $DownloadedFileFullPath = join-path -path $TempDir -ChildPath $($Item.name)
                    $FinalFileName = ($NamingMask + "-" + $Count + "." + $ActivityFileType)
                    $FinalFileNameTempFullPath = join-path -path $TempDir -ChildPath $FinalFileName
                    if (Test-Path $FinalFileNameTempFullPath) {
                        #Allways overwrite temp files
                        $null = Remove-Item $FinalFileNameTempFullPath -Force
                    }
                    Rename-Item $DownloadedFileFullPath -NewName $FinalFileNameTempFullPath -Force
                    if ($Overwrite -eq "Yes") {
                        Move-Item $FinalFileNameTempFullPath -Destination $Destination -Force
                    }
                    else {
                        if (Test-Path (join-path -path $Destination -childpath $FinalFileName)) {
                            write-warning "WARNING - Skipping file $FinalFileName because it allready exists in the destination directory."
                        }
                        else {
                            Move-Item $FinalFileNameTempFullPath -Destination $Destination
                        }
                    }

                }
            }

            elseif ($zip.items().count -eq 1) {
                write-host "INFO - Trying to create file $NamingMask.$ActivityFileType"
                foreach ($Item in $ZIP.items()) {
                    $Shell.Namespace(“$TempDir”).copyhere($zip.items(), 0x14)
                    $DownloadedFileFullPath = join-path -path $TempDir -ChildPath $($Item.name)
                    $FinalFileName = ($NamingMask + "." + $ActivityFileType)
                    $FinalFileNameTempFullPath = join-path -path $TempDir -ChildPath $FinalFileName
                    if (Test-Path $FinalFileNameTempFullPath) {
                        #Allways overwrite temp files
                        $null = Remove-Item $FinalFileNameTempFullPath -Force
                    }
                    Rename-Item $DownloadedFileFullPath -NewName $FinalFileNameTempFullPath -Force
                    if ($Overwrite -eq "Yes") {
                        Move-Item $FinalFileNameTempFullPath -Destination $Destination -Force
                    }
                    else {
                        if (Test-Path (join-path -path $Destination -ChildPath $FinalFileName)) {
                            write-warning "WARNING - Skipping file $FinalFileName because it allready exists in the destination directory."
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
            write-host "INFO - Trying to create file $NamingMask.$ActivityFileType"
            $FinalFileName = ($NamingMask + "." + $ActivityFileType)
            $FinalFileNameTempFullPath = Join-Path -path $TempDir -ChildPath $FinalFileName
            Rename-Item $OutputFileFullPath -NewName $FinalFileNameTempFullPath
            if ($Overwrite -eq "Yes") {
                Move-Item $FinalFileNameTempFullPath -Destination $Destination -Force
            }
            else {
                if (Test-Path (join-path -path $Destination -ChildPath $FinalFileName)) {
                    write-warning "WARNING - Skipping file $FinalFileName because it allready exists in the destination directory."
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
    write-error "ERROR - An unexpected error occurred. See errordetails below:`n$($error[0])"
    break
}

#Finally write the hidden delta cookie for later use
if (($Activities[0].activity.activityId).length -gt 0) {
    if (!(Test-Path $CookieFileFullPath)) {$null = New-Item $CookieFileFullPath -ItemType File -Force}
    $NewestActivity = ($Activities[0].activity.activityId)
    try {

        $CookieFileFullPath = join-path -path $Destination -ChildPath $CookieFilename
        (Get-Item $CookieFileFullPath -Force).Attributes = "Normal"
        $NewestActivity | Out-File $CookieFileFullPath -Force
        (Get-Item $CookieFileFullPath -Force).Attributes = "Hidden"
        write-host "INFO - Finished exporting $ActivityExportedCount activities from Garmin Connect. Delta file successfully stored."
    }
    catch {
        write-error "ERROR - Unable the write the delta file for later use, see error details below:`n$($error[0])"
    }
}
else {write-warning "No new activities found." }