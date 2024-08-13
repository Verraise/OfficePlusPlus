$MainPath = Read-host "请输入项目路径"
$SourceFolder = "$MainPath\P3 docx2zip"
$DestinationFolder = "$MainPath\P4 unzipped"

if (-not (Test-Path -Path $DestinationFolder)) {
    New-Item -ItemType Directory -Path $DestinationFolder | Out-Null
}

$ZipFiles = Get-ChildItem -Path $SourceFolder -Filter *.zip -Recurse

foreach ($ZipFile in $ZipFiles) {
    # 获取 zip 文件相对于源文件夹的路径
    $RelativePath = $ZipFile.FullName.Substring($SourceFolder.Length + 1)
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
