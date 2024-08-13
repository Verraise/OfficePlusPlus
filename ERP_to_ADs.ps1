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
	$path = ReadInput_Text -message "请输入路径"
# $csvFile= read-host "请输入csv路径"
# 总之这里先写死
	$csvFile = "D:\TUVHD-データベース\2 データベース\2-3 テンプレート\2-3-2 证书\0_ERP转ADs\dic.csv"
	$erpFolder = "ERP"
	$adsFolder = "ADs"

#检测ADs文件夹如不存在则新建
if (-not (Test-Path -Path "${path}\$adsFolder")) {
        New-Item -Path "${path}\$adsFolder" -ItemType Directory
		cls
		Write-Host "未检测到ADs文件夹，即将新建" -foregroundcolor yellow
        Write-Host "已创建ADs文件夹: "${path}\$adsFolder"" -foregroundcolor cyan
    } else {
		cls
        Write-Host "ADs文件夹已存在: "${path}\$adsFolder"" -foregroundcolor red
		Write-Host "按回车会覆盖已经存在的ADs证书模板" -foregroundcolor red
		Write-Host "如果要保留，直接关掉即可" -foregroundcolor red
		pause
    }	


#杀掉现有的WORD进程，如果不需要，则注释掉这一行
	Get-Process | Where-Object { $_.ProcessName -eq "WINWORD" } | Stop-Process -Force
	Write-Host "初始化完毕" -foregroundcolor YELLOW
	
# 清空ADs文件夹中的所有文件
	Get-ChildItem -Path (Join-Path -Path $path -ChildPath $adsFolder) | Remove-Item -Force

# saveas方法在另存为时大概率卡死，不知道是不是pwsh的bug，总之先复制过去绕个弯子吧
	copy-item "${path}\${erpFolder}\*" -destination "${path}\$adsFolder" -Recurse -PassThru

# 读取CSV字典
	$csvData = Import-Csv $csvFile

# 获取ADs里面的docx文件
	$ADsDocxFiles = Get-ChildItem -Path (Join-Path -Path $path -ChildPath $adsFolder) -Filter "*.docx"

# 处理ADs文件夹中的每个docx文件
foreach ($docxFile in $ADsDocxFiles) {
	Write-Host "正在启动Word……" -foregroundcolor YELLOW
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false
	
	Write-Host "正在读取文件……"
    $doc = $wordApp.Documents.Open($docxFile.FullName, $false, $false)
	Write-Host "$docxFile 读取成功" -foregroundcolor cyan

# 引用字典
    foreach ($row in $csvData) {
        $findString = $row.str1
        $replaceString = $row.str2

    # 执行替换操作
        $findReplace = $doc.Content.Find.Execute($findString, $true, $true, $false, $false, $false, $true, 1, $false, $replaceString, 2)
        if ($findReplace) {
            Write-Host "替换成功：'$findString' 已替换为 '$replaceString'" -foregroundcolor cyan
        } else {
            Write-Host "未找到匹配项：'$findString'"
        }
    }

# 保存
    Write-Host "正在保存……"
    $doc.Save()
    Write-Host "$docxFile 保存成功" -foregroundcolor cyan
	Write-Host "正在关闭……"
    $doc.Close()
	Write-Host "$docxFile 关闭成功" -foregroundcolor cyan
	Write-Host "正在退出Word……"
    $wordApp.Quit()
	Write-Host "Word退出成功" -foregroundcolor cyan
	
}

    # 释放对象
	Write-Host "正在清除内存中的com对象（1/2）……"
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
	Write-Host "正在清除内存中的com对象（2/2）……"
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null