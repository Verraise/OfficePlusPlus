$dir1 = "C:\Users\34913\Desktop\TUVHD\1 プロジェクト\1-2 证书\1-2-1 产品认证"
$dir2 = "C:\Users\34913\Desktop\TUVHD\1 プロジェクト\1-2 证书\1-2-2 服务认证"
$dir3 = "C:\Users\34913\Desktop\TUVHD\1 プロジェクト\1-2 证书\1-2-3 管理体系认证"
$outTxt = "C:\Users\34913\Desktop\ツール\スクリプト\Test-Path\dirName.csv"

$subfolders1 = Get-ChildItem -Path $dir1 -Directory | Select-Object Name
$subfolders2 = Get-ChildItem -Path $dir2 -Directory | Select-Object Name
$subfolders3 = Get-ChildItem -Path $dir3 -Directory | Select-Object Name
$allSubfolders = $subfolders1 + $subfolders2 + $subfolders3

$allSubfolders | Out-File -FilePath $outTxt -Force

Write-Host "目录文件列表已经输出到了"$outTxt

invoke-item $outTxt