//%attributes = {"invisible":true,"preemptive":"capable"}
/*
generate Settings/buildApp.4DSettings
the build target path, which is ../../../
must be specified using platform full path
the generated file must not be put under source control
as it automatically contains license numbers!
*/
var $packageFolder : 4D:C1709.Folder
$packageFolder:=Folder:C1567(fk database folder:K87:14)

var $databaseFolderPath : Text
$databaseFolderPath:=$packageFolder.platformPath

var $rootFolderPath : 4D:C1709.Folder
$rootFolderPath:=Folder:C1567($databaseFolderPath; fk platform path:K87:2).parent.parent

var $buildAppTemplateFile : 4D:C1709.File
$buildAppTemplateFile:=$packageFolder.file("buildApp.4DSettings")

If ($buildAppTemplateFile.exists)
	var $template; $buildAppSettings : Text
	$template:=$buildAppTemplateFile.getText()
	PROCESS 4D TAGS:C816($template; $buildAppSettings; $packageFolder.name; $rootFolderPath.platformPath)
	$settingsFolder:=$packageFolder.folder("Settings")
	$settingsFolder.create()
	var $buildAppSettingsFile : 4D:C1709.File
	$buildAppSettingsFile:=$settingsFolder.file($buildAppTemplateFile.fullName)
	$buildAppSettingsFile.setText($buildAppSettings)
End if 
