$date  =  "{0:yyyy-MM-dd}"  -f  ( get-date )
$pathtest = "F:\BAK\"+"$date"
$result = Test-Path $pathtest
if ($result -eq "True")
{Write-output "$pathtest already exists.Backup FAILED!"
PAUSE
}
	$path  =  New-Item -path "D:\BAK\" -name  $date  -ItemType directory
	Write-Output "Backup started, please wait..."
	copy-item "C:\Users\34913\Desktop\TUVHD\*" -destination $path -Recurse -PassThru
	function Read-MessageBoxDialog
	{
		$notice = new-object -comobject wscript.shell
		$notice.popup("Done.")
	}
	Read-MessageBoxDialog