# Garmin Connect Activity Export

## Description
This script lets you backup your activities stored in Garmin Connect. It supports backups in de the following formats:
- FIT (default)
- GPX
- TCX
- KML

It supports two download options:
- Delta, exporting only new activities (default)
- Full, exporting all activities available.

The scripts does the following:
 - Downloads activity files from garmin in FIT, TCX or GPX format.
 - Supports delta download

## Commandline options
The following commandline options are available:<br>
- ActivityFileType (Optional) - Choose the activitiy file type to export in. FIT is the default value.
	- FIT (default)
	- TCX
	- GPX
	- KML
- DownloadOption (Optional) - Choose to download only new activities or all activities. At the first time all activities are downloaded. Depends on a cookie file placed in the destination folder.
	- New (Default)
	- All
- Destination (Mandatory) - The location on your device where the files need to be exported to.
- Username (Mandatory) - Your Garmin Connect username.
- Password (Mandatory) - Your Garmin Connect account password.
- Overwrite (Optional) - Overwrite existing files in the destination location.
	- Yes
	- No (Default)

You can change the default options the following XML file:
- GCUserSettings: Containing specific settings for your situation;

## Examples
Download only new activities in the FIT format:<br>
```& C:\Scripts\GCActivityExport.ps1 -Destination "c:\GarminActivities" -UserName "<Your Garmin Connect Username>" -Password "<Your Garmin Connect Password>"```

Download all activities in TCX format overwriting all files in the destination:<br>
```& C:\Scripts\GCActivityExport.ps1 -Destination "c:\GarminActivities" -UserName "<Your Garmin Connect Username>" -Password "<Your Garmin Connect Password>" -ActivityFileType TCX -DownloadOption All -Overwrite Yes```

## Version history<br>
1.0   - Initial version<br>
1.1   - Fix: Garmin now expects parameters in the SSO url<br>
        Update: Added settings support in separate XML files<br>
1.2   - Update support for new Garmin activity feed<br>
1.3   - Update to support the new Garmin Signin URL<br>
1.3.1 - Fix due error 402 Payment required error when retrieving activity list<br>
1.3.2 - Remove default destination folder, because it causes to ignore the GCUserSettings.xml destination<br>
        Remove empty strings for Username and Password<br>
        Format layout<br>
1.4   - Update authentication process to use an OAuth token (thanks to @ShiveringPenguin)<br>
        Add support for KML files<br>
        Updated user-agent<br>
        Remove Garmin Connect Actvity Export -Program Settings.xml dependency<br>
        Layout and output improvements<br>
        Report when rate limited (or any other error from Garmin)<br>
## Credits
Credits to Kyle Krafka (https://github.com/kjkjava/) for delivering a great example for this script.
