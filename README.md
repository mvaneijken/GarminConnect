# Garmin Connect Activity Export

Description:
This script lets you backup your activities stored in Garmin Connect. It supports backups in de the following formats:
- FIT (default)
- GPX
- TCX

It supports two download options:
- Delta, exporting only new activities (default)
- Full, exporting all activities available. 


Version history:
 1.0 - Initial version 
 1.1 - Fix: Garmin now expects parameters in the SSO url
       Update: Added settings support in separate XML files

The scripts does the following:
 - Downloads activity files from garmin in FIT, TCX or GPX format.
 - Supports delta download

The following commandline options are available:
-ActivityFileType (Optional) - Choose the activitiy file type to export in. FIT is the default value.
	- FIT (default)
	- TCX
	- GPX
-DownloadOption (Optional) - Choose to download only new activities or all activities. At the first time all activities are downloaded. Depends on a cookie file placed in the destination folder. 
	- New (Default)
	- All 
-Destination (Mandatory) - The location on your device where the files need to be exported to.
-Username (Mandatory) - Your Garmin Connect username.
-Password (Mandatory) - Your Garmin Connect account password.
-Overwrite (Optional) - Overwrite existing files in the destination location.
	- Yes
	- No (Default)

You can change the default options the following XML files:
- GCUserSettings: Containing specific settings for your situation;
- GCProgramSettings: Containing configurations regarding Garmin Connect. These are only needed to be changed when Garmin changes their URLs.

Examples:

Download only new activities in the FIT format
	GCActivityExport.ps1 -Destination "c:\GarminActivities" -UserName <Your Garmin Connect Username> 

Download all activities in TCX format overwriting all files in the destination
	GCActivityExport.ps1 -Destination "c:\GarminActivities" -UserName <Your Garmin Connect Username> -ActivityFileType TCX -DownloadOption All -Overwrite Yes

Credits to Kyle Krafka (https://github.com/kjkjava/) for delivering a great example for this script. 
