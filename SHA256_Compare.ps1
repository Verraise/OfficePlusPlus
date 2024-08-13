$path = read-host "input directory"
cd $path
$file1 = read-host "file 1"
$file2 = read-host "file 2"

$md5File1 = Get-FileHash -Path $file1 | Select-Object -ExpandProperty Hash
$md5File2 = Get-FileHash -Path $file2 | Select-Object -ExpandProperty Hash

if ($md5File1 -eq $md5File2) {
    Write-Host "文件的MD5码相同" -foregroundColor green
} else {
    Write-Host "文件的MD5码不同" -foregroundColor red
    Write-Host "文件1的MD5码：$md5File1"
    Write-Host "文件2的MD5码：$md5File2"
}

cmd -command "pause"