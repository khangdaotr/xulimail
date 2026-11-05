# Tạo tác vụ chạy main.py mỗi 5 phút (ẩn cửa sổ)
$py = (Get-Command pythonw).Source
$script = Join-Path $PSScriptRoot "main.py"
schtasks /create /f /tn "XuLiMail" /sc minute /mo 5 /tr "`"$py`" `"$script`""
# Chạy bù ngay khi máy thức dậy; chạy quá 10 phút (treo) thì tự dừng để lượt sau không bị chặn
$s = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 10) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Set-ScheduledTask -TaskName "XuLiMail" -Settings $s | Out-Null
Write-Host "Đã tạo tác vụ XuLiMail. Xóa bằng: schtasks /delete /tn XuLiMail /f"
