$pathfillstr = '\1'
$count = 0
$UnzippedFolder = Read-Host "请输入即将被替换的文件夹路径："
$RepositoryPath = Read-Host "请输入这一轮要替换的新底板图片文件夹的路径："

# 获取旧底板文件夹中所有大于150KB的.png、.jpg、.jpeg文件
$FilesToReplace = Get-ChildItem -Path $UnzippedFolder -Include *.png,*.jpg,*.jpeg -Recurse | Where-Object { $_.Length -gt 150KB }

cls

foreach ($File in $FilesToReplace) {
	$TargetFilePath = $File.fullname
	remove-item -path $TargetFilePath -force
    # 查找对应的新底板
    $RepositoryFilePath = "$RepositoryPath" + $pathfillstr + $File.Extension	
    if (Test-Path $RepositoryFilePath) {
        copy-item -path $RepositoryFilePath -destination $TargetFilePath -force
		$count++
		write-host "已替换 $count 个目标" -foregroundcolor green
    }
}

Write-Host "内容替换完成。"

pause