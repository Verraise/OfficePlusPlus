# 设置要处理的文件夹路径，包含需要合并的 TS 文件
$folderPath = Read-host "Input dir"

# 获取文件夹中的所有 TS 文件，并按照文件名排序
$tsFiles = Get-ChildItem -Path $folderPath -Filter *.ts | Sort-Object Name

# 创建一个新的 total.ts 文件
$totalFilePath = Join-Path -Path $folderPath -ChildPath "total.ts"
New-Item -ItemType File -Path $totalFilePath -Force | Out-Null

# 遍历每个 TS 文件，逐个将其内容追加到 total.ts 文件中
foreach ($file in $tsFiles) {
    Add-Content -Path $totalFilePath -Value (Get-Content $file.FullName)
}

Write-Host "All TS files have been merged into total.ts"