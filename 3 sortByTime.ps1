# 设置要处理的文件夹路径
$folderPath = Read-Host "Input dir"


<#
# 获取文件夹中的所有 PNG 文件，并按照创建时间进行排序
$pngFiles = Get-ChildItem -Path $folderPath -Filter *.png | Sort-Object LastWriteTime

# 遍历每个 PNG 文件，重命名为顺序递增的文件名
for ($i = 0; $i -lt $pngFiles.Count; $i++) {
    $newFileName = '{0:D4}.ts' -f ($i + 1)
    $newFilePath = Join-Path -Path $folderPath -ChildPath $newFileName
    $pngFiles[$i] | Rename-Item -NewName $newFileName -Force
    Write-Host "Renamed $($pngFiles[$i].Name) to $newFileName"
}

#>


# 获取文件夹中的所有文件，并按照创建时间进行排序
$allFiles = Get-ChildItem -Path $folderPath | Sort-Object LastWriteTime

# 遍历每个 PNG 文件，重命名为顺序递增的文件名
for ($i = 0; $i -lt $allFiles.Count; $i++) {
    $newFileName = '{0:D4}.ts' -f ($i + 1)
    $newFilePath = Join-Path -Path $folderPath -ChildPath $newFileName
    $allFiles[$i] | Rename-Item -NewName $newFileName -Force
    Write-Host "Renamed $($allFiles[$i].Name) to $newFileName"
}
