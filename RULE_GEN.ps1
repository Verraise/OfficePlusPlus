#杀掉现有的WORD进程，如果不需要，则注释掉这一行
	Get-Process | Where-Object { $_.ProcessName -eq "WINWORD" } | Stop-Process -Force

#输入值空值校验函数
function ReadInput_Text {
    param([string]$message, [string]$defaultValue)
    while($true) {
        $input = Read-Host $message
        if(![String]::IsNullOrWhiteSpace($input)) {
            return $input.Trim()
        }
        if($defaultValue) {
            return $defaultValue
        }
    }
}

# 设置路径和文件名
	$path = ReadInput_Text -message "请输入规则路径"
	$csvFile = "dic.csv"
	$erpFolder = "ERP"
	$adsFolder = "ADs"

# 清空ADs文件夹中的所有文件
	Get-ChildItem -Path (Join-Path -Path $path -ChildPath $adsFolder) | Remove-Item -Force

# 读取CSV字典
	$csvData = Import-Csv -Path (Join-Path -Path $path -ChildPath $csvFile)

#构建规则文件名称
	$ruleNum_raw = ($csvData | Where-Object { $_.str1 -eq '${规则号}' }).str2
	$ruleNum0 = $ruleNum_raw -replace "/", ""
	$ruleNum = $ruleNum0 -replace "GZF", "GZ F"
	$ruleName_CN = ($csvData | Where-Object { $_.str1 -eq '${全名}' }).str2
	$ruleName_Full = "${ruleNum} ${ruleName_CN}认证规则.docx"
	Write-Host "构建规则名完毕：$ruleName_Full" -foregroundcolor cyan

# saveas方法在另存为时大概率卡死，不知道是不是pwsh的bug，总之先复制过去绕个弯子吧
	cls
	Write-Host "正在复制副本……"
	copy-item "${path}\${erpFolder}\*" -destination "${path}\$adsFolder\$ruleName_Full" -Recurse -PassThru
	cls
	Write-Host "复制完成！" -foregroundcolor cyan
	
# 获取ADs里面的docx文件
	Write-Host "正在获取文件列表……"
	$ADsDocxFiles = Get-ChildItem -Path (Join-Path -Path $path -ChildPath $adsFolder) -Filter "*.docx"
	Write-Host "文件列表获取完毕！" -foregroundcolor cyan
# 处理ADs文件夹中的每个docx文件
	Write-Host "正在读取文件……"
	foreach ($docxFile in $ADsDocxFiles) {
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false
    $doc = $wordApp.Documents.Open($docxFile.FullName, $false, $false)
	Write-Host "读取完毕，准备替换！" -foregroundcolor cyan
	
	# 引用字典
		foreach ($row in $csvData) {
			$findString = $row.str1
			$replaceString = $row.str2

			$storyRanges = $doc.StoryRanges
			foreach ($storyRange in $storyRanges) {
				$findReplaceContent = $storyRange.Find.Execute($findString, $true, $true, $false, $false, $false, $true, 1, $false, $replaceString, 2)
				if ($findReplaceContent) {
				$execCount ++
				Write-Host "第 $execCount 个关键词已替换成功：'$findString' 已替换为 '$replaceString'" -foregroundcolor cyan
				} else {
				#Write-Host "未找到匹配项：'$findString'"
				}
			}
		}
	
	cls
	Write-Host "已成功替换 $execCount 个关键词" -foregroundcolor cyan
	
	#更新目录
	Write-Host "正在更新目录"
	$doc.TablesOfContents | ForEach-Object { $_.Update() }
	Write-Host "已更新目录。" -foregroundcolor cyan

	# 保存
		Write-Host "正在保存……"
		$doc.Save()
		Write-Host "$modifiedDocxPath 保存成功" -foregroundcolor cyan

		$doc.Close()
		$wordApp.Quit()


    # 释放对象
	Write-Host "正在清除内存中的com对象……"
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null
	Write-Host "清除完成" -foregroundcolor cyan
	}
	
#打开替换后的文件
Write-Host "正在打开替换后的规则文件……"
ii "${path}\$adsFolder\$ruleName_Full"