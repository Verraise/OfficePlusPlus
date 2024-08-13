$count_file = 0
# 输入日期和时间
$desiredDateTime = Read-Host "输入日期和时间 (YYYY-MM-DD HH:MM:SS)"

# 将输入的日期时间字符串转换为 DateTime 对象
$dateTime = [DateTime]::ParseExact($desiredDateTime, "yyyy-MM-dd HH:mm:ss", $null)

# 获取目标文件夹中的所有文件和文件夹
$destinationItems = Get-ChildItem -Path $destinationFolder -Recurse

foreach ($destinationItem in $destinationItems) {
    # 修改文件/文件夹的创建时间和修改时间为指定值
    $destinationItem.CreationTime = $dateTime
    $destinationItem.LastWriteTime = $dateTime

Write-Host "属性信息已清除完成。"


	write-host "已修改 $count_file 个文件时间信息" -foregroundcolor green
	$count_file++
}

Write-Host "操作完成。"

PAUSE

ii $destinationFolder