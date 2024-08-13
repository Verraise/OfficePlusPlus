# 设置路径和文件名
	$path = read-host "请输入路径"
# $csvFile= read-host "请输入csv路径"
# 总之这里先写死
	#$csvFile = "C:\Users\34913\Desktop\TUVHD\TMP\新しいフォルダー\ERP证书字段替换\dic.csv"
	$csvFile = read-host "当前是外链csv模式，请输入外链csv路径"
	$erpFolder = "ERP"
	$EXPFolder = "EXP"

#检测EXP文件夹如不存在则新建
if (-not (Test-Path -Path "${path}\$EXPFolder")) {
        New-Item -Path "${path}\$EXPFolder" -ItemType Directory
		cls
		Write-Host "未检测到EXP文件夹，即将新建" -foregroundcolor yellow
        Write-Host "已创建EXP文件夹: "${path}\$EXPFolder"" -foregroundcolor cyan
    } else {
		cls
        Write-Host "EXP文件夹已存在: "${path}\$EXPFolder"" -foregroundcolor red
		Write-Host "按回车会覆盖已经存在的EXP证书模板" -foregroundcolor red
		Write-Host "如果要保留，直接关掉即可" -foregroundcolor red
		pause
    }	


#杀掉现有的WORD进程，如果不需要，则注释掉这一行
	Get-Process | Where-Object { $_.ProcessName -eq "WINWORD" } | Stop-Process -Force
	Write-Host "初始化完毕" -foregroundcolor YELLOW
	
# 清空EXP文件夹中的所有文件
	Get-ChildItem -Path (Join-Path -Path $path -ChildPath $EXPFolder) | Remove-Item -Force

# saveas方法在另存为时大概率卡死，不知道是不是pwsh的bug，总之先复制过去绕个弯子吧
	copy-item "${path}\${erpFolder}\*" -destination "${path}\$EXPFolder" -Recurse -PassThru

# 读取CSV字典
	$csvData = Import-Csv $csvFile

# 获取EXP里面的docx文件
	$EXPDocxFiles = Get-ChildItem -Path (Join-Path -Path $path -ChildPath $EXPFolder) -Filter "*.docx" -recurse

# 处理EXP文件夹中的每个docx文件
foreach ($docxFile in $EXPDocxFiles) {
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