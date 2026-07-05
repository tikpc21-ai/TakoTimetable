# TakoTimetable Local Network Web Server
# Allows teachers and students to access the schedule from their phones/devices on the same Wi-Fi.

$port = 8080
$localIP = (Get-NetIPAddress | Where-Object { $_.AddressState -eq "Preferred" -and $_.ValidLifetime -lt [TimeSpan]::MaxValue -and $_.IPAddress -like "192.168.*" }).IPAddress

if (-not $localIP) {
    $localIP = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -ne "127.0.0.1" } | Select-Object -First 1).IPAddress
}

if (-not $localIP) {
    $localIP = "localhost"
}

$urlLocal = "http://localhost:$port/"
$urlNetwork = "http://$($localIP):$port/"

# Create HTTP Listener
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($urlLocal)
if ($localIP -ne "localhost") {
    $listener.Prefixes.Add($urlNetwork)
}

Write-Host "==========================================================" -ForegroundColor Green
Write-Host "  TakoTimetable Local Web Server is Running!" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
Write-Host " Host Machine: $urlLocal" -ForegroundColor Yellow
Write-Host " Local Network (Wi-Fi): $urlNetwork" -ForegroundColor Cyan
Write-Host "----------------------------------------------------------"
Write-Host " Connect your smartphone or other PCs to the same Wi-Fi"
Write-Host " and open the Local Network URL above to access the schedules."
Write-Host " Press [Ctrl + C] or close this window to STOP the server."
Write-Host "=========================================================="

# Auto-open browser on local host
Start-Process $urlLocal

try {
    $listener.Start()
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        
        $reqPath = $request.Url.LocalPath
        if ($reqPath -eq "/") {
            $reqPath = "/index.html"
        }
        
        $localFilePath = Join-Path $pwd $reqPath.Substring(1)
        
        if (Test-Path $localFilePath -PathType Leaf) {
            $bytes = [System.IO.File]::ReadAllBytes($localFilePath)
            
            # Set content-type
            $ext = [System.IO.Path]::GetExtension($localFilePath).ToLower()
            switch ($ext) {
                ".html" { $response.ContentType = "text/html; charset=utf-8" }
                ".css"  { $response.ContentType = "text/css; charset=utf-8" }
                ".js"   { $response.ContentType = "application/javascript; charset=utf-8" }
                ".json" { $response.ContentType = "application/json; charset=utf-8" }
                ".png"  { $response.ContentType = "image/png" }
                ".jpg"  { $response.ContentType = "image/jpeg" }
                default { $response.ContentType = "application/octet-stream" }
            }
            
            $response.ContentLength64 = $bytes.Length
            $response.OutputStream.Write($bytes, 0, $bytes.Length)
        } else {
            $response.StatusCode = 404
            $errBytes = [System.Text.Encoding]::UTF8.GetBytes("404 File Not Found")
            $response.OutputStream.Write($errBytes, 0, $errBytes.Length)
        }
        $response.OutputStream.Close()
    }
}
catch {
    Write-Host "Server stopped: $_" -ForegroundColor Red
}
finally {
    $listener.Stop()
    $listener.Close()
}
