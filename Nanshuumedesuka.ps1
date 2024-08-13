$weekNum = Get-Date -UFormat %V
$weekNum


function Read-MessageBoxDialog

{

	$PopUpWin = new-object -comobject wscript.shell

	$PopUpWin.popup("现在是第"+$weekNum+"周")

}

Read-MessageBoxDialog