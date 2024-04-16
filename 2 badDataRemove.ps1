# 获取文件夹中的所有 PNG 文件
$folderPath = "C:\Users\Administrator\Desktop\222"
$pngFiles = Get-ChildItem -Path $folderPath -Filter *.ts

# 遍历每个 PNG 文件
foreach ($file in $pngFiles) {
    # 读取文件的二进制数据
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    
    # 查找不良信息的结尾
    $indexOfMaliciousDataEnd = -1
    for ($i = 0; $i -lt $bytes.Length - 3; $i++) {
        if ($bytes[$i] -eq 127 -and $bytes[$i + 1] -eq 255 -and $bytes[$i + 2] -eq 217 -and $bytes[$i + 3] -eq 71) {
            $indexOfMaliciousDataEnd = $i + 3
            break
        }
    }

$indexOfMaliciousDataEnd

    if ($indexOfMaliciousDataEnd -ne -1) {
        # 删除不良信息的起始部分及其之前的所有数据
        $bytes = $bytes[$indexOfMaliciousDataEnd..$bytes.Length]
        
        # 保存修复后的数据到同名文件中
        [System.IO.File]::WriteAllBytes($file.FullName, $bytes)
        Write-Host "Removed malicious data from $($file.Name)"
    } else {
        Write-Host "No malicious data found in $($file.Name)"
    }
}