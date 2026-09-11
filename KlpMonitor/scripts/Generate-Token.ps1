$bytes = New-Object byte[] 32
$generator = [System.Security.Cryptography.RandomNumberGenerator]::Create()
try {
    $generator.GetBytes($bytes)
} finally {
    $generator.Dispose()
}
$token = ([System.BitConverter]::ToString($bytes)).Replace('-', '').ToLowerInvariant()
Write-Output $token
