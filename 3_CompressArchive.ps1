# 1. 筛选文件夹
$mainFolder = Read-Host "请输入P6文件夹路径"
$zipFolders = Get-ChildItem -Path $mainFolder -Filter "*.zip" -Directory -Recurse

# 2. 修改文件夹名称
foreach ($folder in $zipFolders) {
    $newName = $folder.Name + "1"
    Rename-Item -Path $folder.FullName -NewName $newName
	$countRenameFolder++
	Write-host "已调整 $countRenameFolder 个文件夹名称。" -foregroundcolor green
}

#3、还原文件夹为zip文件
foreach ($folder in $zipFolders) {
    $sourcePath = $folder.FullName + "1"
	$parentDirectory = Split-Path -Path $folder.FullName -Parent
	$destinationPath = Join-Path -Path $parentDirectory -ChildPath ($folder.Name -replace '\.zip1$', '.zip')
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory($sourcePath, $destinationPath)
	$countArchived++
	Write-host "已压缩 $countArchived 个文件夹。" -foregroundcolor green
}


# 4. 删除所有筛选出来的文件夹
$zip1Folders = Get-ChildItem -Path $mainFolder -Filter "*.zip1" -Directory -Recurse
foreach ($folder in $zip1Folders) {
    Remove-Item -Path $folder.FullName -Force -Recurse
		$countDel++
	Write-host "已删除 $countDel 个zip1中间文件。" -foregroundcolor green
}

# 5、把所有生成的".zip"文件的后缀，全部强制修改为"docx"。
$generatedZipFiles = Get-ChildItem -Path $mainFolder -Filter "*.zip" -File -Recurse
foreach ($file in $generatedZipFiles) {
    $newFileName = $file.FullName -replace '\.zip$', '.docx'
    Rename-Item -Path $file.FullName -NewName $newFileName
	$countRenameDocx++
	Write-host "已生成 $countRenameDocx 个docx。" -foregroundcolor green
}
pause