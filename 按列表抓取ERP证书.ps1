# 设置文件路径
$csvFilePath = "C:\Users\34913\Desktop\1.csv"
$oldFolderPath = "C:\Users\34913\Desktop\TUVHD\1 プロジェクト\1-2 证书"
$newFolderPath = "C:\Users\34913\Desktop\TUVHD\TMP\ERPCERTBAK\P5"

# 读取CSV文件中的内容
$csvContent = Import-Csv $csvFilePath

# 遍历"old"文件夹及其子文件夹
$filesToCopy = @()
foreach ($csvEntry in $csvContent) {
    $fileName = $csvEntry.名前
    $files = Get-ChildItem -Path $oldFolderPath -Recurse -Filter $fileName -File | Select-Object -ExpandProperty FullName
    if ($files) {
        $filesToCopy += $files
    }
}

# 将找到的文件复制到"new"文件夹中
foreach ($fileToCopy in $filesToCopy) {
    $destinationPath = Join-Path -Path $newFolderPath -ChildPath (Split-Path -Path $fileToCopy -Leaf)
    Copy-Item -Path $fileToCopy -Destination $destinationPath -Force
}
