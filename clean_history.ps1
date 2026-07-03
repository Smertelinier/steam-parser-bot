param([string]$token, [string]$placeholder)
$files = @('bot.py', 'Dockerfile', 'requirements.txt', 'users.json', '.env', '.env.example')
foreach ($f in $files) {
    $path = Join-Path (Get-Location) $f
    if (Test-Path $path) {
        $c = Get-Content $path -Raw -ErrorAction SilentlyContinue
        if ($c) {
            $c = $c -replace $token, $placeholder
            [System.IO.File]::WriteAllText($path, $c, [System.Text.UTF8Encoding]::new($false))
        }
    }
}
