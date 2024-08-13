$dir = read-host "请输入要导出的目录路径"
$outTxt = "C:\Users\34913\Desktop\TUVHD\TMP\dirName.txt"

Get-ChildItem -Path $dir | ForEach-Object { $_.Name } | Out-File -FilePath $outTxt

Write-Host "目录文件列表已经输出到了"$outTxt
invoke-item $outTxt