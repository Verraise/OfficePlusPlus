$txtFilePath = read-host "inputTXT"

$destinationFolder = "C:\Users\Administrator\Desktop\TSDL\OUTPUT"

Get-Content $txtFilePath | ForEach-Object {
    $url = $_
    $fileName = Split-Path -Leaf $url
	
	$fileName = $fileName -replace "\?container=cmaf", ""
	
    $outputPath = Join-Path -Path $destinationFolder -ChildPath $fileName 
	

	
    Invoke-WebRequest -Uri $url -OutFile $outputPath
    Write-Host "Downloaded: $fileName"
}


	
	
	
	