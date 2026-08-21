$ping = Test-Connection -ComputerName 1.1.1.1 -Count 4
Write-Host "--- PING TEST (Latency) ---"
$ping | Select-Object Address, ResponseTime | Format-Table -AutoSize

Write-Host "--- DOWNLOAD SPEED TEST ---"
try {
    $time = Measure-Command { 
        Invoke-WebRequest -Uri "https://speed.cloudflare.com/__down?bytes=10000000" -OutFile "$env:TEMP\test_speed.tmp" -UseBasicParsing 
    }
    $seconds = $time.TotalSeconds
    $mbps = (10 * 8) / $seconds
    Write-Host "Time to download 10MB: $seconds seconds"
    Write-Host "Estimated Speed: $mbps Mbps"
} catch {
    Write-Host "Failed to download test file. Blocked by firewall?"
    Write-Host $_.Exception.Message
}
