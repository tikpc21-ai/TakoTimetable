$path = "index.html"
$content = [System.IO.File]::ReadAllText($path)
if (-not $content.Contains('rel="manifest"')) {
    $replacement = "<link rel=`"manifest`" href=`"manifest.json`">`n<script>`nif(`"serviceWorker`" in navigator) {`nwindow.addEventListener(`"load`", () => {`nnavigator.serviceWorker.register(`"sw.js`");`n});`n}`n</script>`n</head>"
    $content = $content.Replace("</head>", $replacement)
    [System.IO.File]::WriteAllText($path, $content)
    Write-Host "PWA tags injected successfully."
} else {
    Write-Host "PWA tags already exist."
}
