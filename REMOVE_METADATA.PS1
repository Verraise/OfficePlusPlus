$count_file = 0
$sourceFolder = Read-Host "输入原文件夹路径"
$zipTmpFolder = Read-Host "输入存放zip临时文件的文件夹路径"
$destinationFolder = Read-Host "输入目标文件夹路径"
$desiredDateTime = Read-Host "输入日期和时间 (YYYY-MM-DD HH:MM:SS)"

Copy-Item -Path "$sourceFolder\*" -Destination $zipTmpFolder -Recurse

Get-ChildItem -Path $sourceFolder -Recurse | Where-Object { $_.Name -notlike "*.docx*" } | Copy-Item -Destination {Join-Path -Path $destinationFolder -ChildPath $_.FullName.Substring($sourceFolder.length)}


$zipFiles = Get-ChildItem -Path $zipTmpFolder -recurse

foreach ($zipFile in $zipFiles) {
    # 构建新的文件名，将 .docx 替换为 .zip
    $zipFileName = $zipFile.FullName -replace '\.docx$', '.zip'
    # 重命名文件
		try {
		Rename-Item -Path $zipFile.FullName -NewName $zipFileName
	}catch {}
}
#catch了个寂寞

Write-Host "所有 .docx 文件的后缀已经修改为 .zip。"

cls

#unzip

if (-not (Test-Path -Path $DestinationFolder)) {
    New-Item -ItemType Directory -Path $DestinationFolder | Out-Null
}

$ZipFiles = Get-ChildItem -Path $zipTmpFolder -Filter *.zip -Recurse

foreach ($ZipFile in $ZipFiles) {
    # 获取 zip 文件相对于源文件夹的路径
    $RelativePath = $ZipFile.FullName.Substring($zipTmpFolder.Length + 1)
    # 构建解压后的文件夹路径
    $ExtractPath = Join-Path -Path $DestinationFolder -ChildPath $RelativePath

    # 创建解压后的文件夹路径
    if (-not (Test-Path -Path $ExtractPath)) {
        New-Item -ItemType Directory -Path $ExtractPath | Out-Null
    }

    # 解压
    Expand-Archive -Path $ZipFile.FullName -DestinationPath $ExtractPath -Force
}

Write-Host "解压完成。"

Get-ChildItem -Path $destinationFolder -Recurse -Directory -Filter "docProps" | Remove-Item -Recurse -Force


Write-Host "metadata已删除"


$zipFolders = Get-ChildItem -Path $DestinationFolder -Filter "*.zip" -Directory -Recurse

foreach ($folder in $zipFolders) {
    $newName = $folder.Name + "1"
    Rename-Item -Path $folder.FullName -NewName $newName
	$countRenameFolder++
	Write-host "已调整 $countRenameFolder 个文件夹名称。" -foregroundcolor green
}

foreach ($folder in $zipFolders) {
    $sourcePath = $folder.FullName + "1"
	$parentDirectory = Split-Path -Path $folder.FullName -Parent
	$destinationPath = Join-Path -Path $parentDirectory -ChildPath ($folder.Name -replace '\.zip1$', '.zip')
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory($sourcePath, $destinationPath)
	$countArchived++
	Write-host "已压缩 $countArchived 个文件夹。" -foregroundcolor green
}


$zip1Folders = Get-ChildItem -Path $DestinationFolder -Filter "*.zip1" -Directory -Recurse
foreach ($folder in $zip1Folders) {
    Remove-Item -Path $folder.FullName -Force -Recurse
		$countDel++
	Write-host "已删除 $countDel 个zip1中间文件。" -foregroundcolor green
}

$generatedZipFiles = Get-ChildItem -Path $DestinationFolder -Filter "*.zip" -File -Recurse
foreach ($file in $generatedZipFiles) {
    $newFileName = $file.FullName -replace '\.zip$', '.docx'
    Rename-Item -Path $file.FullName -NewName $newFileName
	$countRenameDocx++
	Write-host "已清除 $countRenameDocx 个文件的metadata。" -foregroundcolor green
}

# 将输入的日期时间字符串转换为 DateTime 对象
$dateTime = [DateTime]::ParseExact($desiredDateTime, "yyyy-MM-dd HH:mm:ss", $null)

# 获取目标文件夹中的所有文件和文件夹
$destinationItems = Get-ChildItem -Path $destinationFolder -Recurse

foreach ($destinationItem in $destinationItems) {
    # 修改文件/文件夹的创建时间和修改时间为指定值
    $destinationItem.CreationTime = $dateTime
    $destinationItem.LastWriteTime = $dateTime
	write-host "已修改 $count_file 个文件时间信息" -foregroundcolor green
	$count_file++
}

Write-Host "操作完成。"

PAUSE

ii $destinationFolder