$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
cd $PSScriptRoot
$dirListPath = $PSScriptRoot+"\dirList.txt"

$content = Get-Content -Path $dirListPath -Raw
$cleanedContent = $content -replace """"
$cleanedContent | Set-Content -Path $dirListPath

$results = @()
$results_false = @()
$dirTotal = 0
$dir_true = 0
$dir_false = 0

if ((test-path -path "$PSScriptRoot\Result.csv" ) -and (test-path -path "$PSScriptRoot\Result_false.csv" )){
		rm -path $PSScriptRoot"\Result.csv"
		rm -path $PSScriptRoot"\Result_false.csv"
		Write-host "初始化完成。" -foregroundcolor green
	}
	else{
		write-host "提示：初始化失败！用于记录错误结果的文件不存在！" -foregroundcolor green
		write-host "但这可能说明上次测试是通过的，没有输出任何文件。" -foregroundcolor green
		write-host "也有可能是你自己删的。" -foregroundcolor green
}

$paths = Get-Content -Path $dirListPath | Where-Object { $_ -ne "" }
foreach ($path in $paths) {

    if (Test-Path -Path $path) {
			$results += 1
			$dir_true = $dir_true + 1
			$dirTotal = $dirTotal + 1
		}
		else {
			$results += 0
			$results_false += $path
			$dir_false = $dir_false + 1
			$dirTotal = $dirTotal + 1
		}
}

#cls

Write-host "已测试"$dirTotal"个目录"
Write-host "    其中"$dir_true"个目录正常已确认存档" -foregroundcolor green
Write-host "         "$dir_false"个目录存在异常" -foregroundcolor red

$results | Out-File -FilePath $PSScriptRoot"\Result.csv" -Force
$results_false | Out-File -FilePath $PSScriptRoot"\Result_false.csv" -Force

if ($dir_false -ne 0)	{
	invoke-item Result_false.csv
	}
	else	{
	Write-host "    查验无误，回车退出。" -foregroundcolor green
	pause
}
