$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
cd $PSScriptRoot
$dirListPath = $PSScriptRoot+"\dirList.txt"
$csvFilePath = $PSScriptRoot+"\pdfdocxcheck.csv"

$content = Get-Content -Path $dirListPath -Raw
$cleanedContent = $content -replace """"
$cleanedContent | Set-Content -Path $dirListPath

rm $csvFilePath

# 初始化结果数组
$results = @()

$paths = Get-Content -Path $dirListPath | Where-Object { $_ -ne "" }

foreach ($path in $paths) {
    $result = New-Object PSObject -Property @{
		"PDF Count" = 0
        "DOCX Count" = 0
		"JPG Count" = 0
		"ERPdocx_en Count" = 0
    }

    if ((Test-Path -Path $path -PathType Container) -or (Test-Path -Path ($path + 'ADs') -PathType Container)){
        $docxFiles = Get-ChildItem -Path $path -Filter "*.docx" -Recurse
        $result."DOCX Count" = $docxFiles.Count

        $pdfFiles = Get-ChildItem -Path $path -Filter "*.pdf" -Recurse
        $result."PDF Count" = $pdfFiles.Count
		
		$JPGFiles = Get-ChildItem -Path $path -Filter "*.jpg" -Recurse
        $result."JPG Count" = $JPGFiles.Count
		
		if (Test-Path -Path ($path + 'erp') -PathType Container) {
			$ERPdocx_en = Get-ChildItem -Path ($path + 'erp') -Filter "*en.docx" -Recurse
			$result."ERPdocx_en Count" = $ERPdocx_en.Count
		}
    }
    $results += $result
}

# 导出结果到CSV文件
$results | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8

invoke-item $csvFilePath