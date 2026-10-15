# Execute na maquina SERVIDORA, com o projeto em disco local (nao na unidade M:).
# Para testar SOMENTE com dados ficticios: .\INICIAR_REDE.ps1 -TesteHttp
param(
    [int]$Porta = 8765,
    [string]$Certificado = '',
    [string]$Chave = '',
    [switch]$TesteHttp
)
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
if ($PSScriptRoot.StartsWith('\\')) { throw 'Instale o programa no disco local do servidor, nao em caminho UNC.' }
$driveName = (Split-Path -Qualifier $PSScriptRoot).TrimEnd(':')
$psDrive = Get-PSDrive -Name $driveName -ErrorAction SilentlyContinue
if ($psDrive -and $psDrive.DisplayRoot) { throw 'A pasta atual esta em uma unidade de rede. Copie o projeto para C:\CentralObserver no SERVIDOR.' }
$python = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (!(Test-Path $python)) {
    $cmd = Get-Command py -ErrorAction SilentlyContinue
    if ($cmd) { $python = 'py' }
    else { $python = 'python' }
}
if (!$TesteHttp -and (!$Certificado -or !$Chave)) {
    throw 'HTTPS obrigatorio para dados reais. Informe -Certificado C:\certificados\central.crt -Chave C:\certificados\central.key. Solicite certificados TLS a TI.'
}
$senha = Read-Host 'Defina a senha da Central para esta sessao (minimo 16 caracteres)' -AsSecureString
$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($senha)
try {
    $env:OBS_LAN_PASSWORD = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
} finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
}
if ($env:OBS_LAN_PASSWORD.Length -lt 16) { throw 'A senha deve ter pelo menos 16 caracteres.' }
$env:OBS_LAN_HOST = '0.0.0.0'
$env:OBS_PORT = "$Porta"
$env:OBS_NO_BROWSER = '1'
$env:OBS_TLS_CERT = $Certificado
$env:OBS_TLS_KEY = $Chave
$env:OBS_ALLOW_INSECURE_HTTP = if ($TesteHttp) { '1' } else { '0' }
try {
    Write-Host "Iniciando servidor na porta $Porta. Mantenha esta janela aberta." -ForegroundColor Cyan
    if ($python -eq 'py') { & py -3 server_lan.py }
    else { & $python server_lan.py }
} finally {
    Remove-Item Env:OBS_LAN_PASSWORD -ErrorAction SilentlyContinue
    Remove-Item Env:OBS_TLS_CERT -ErrorAction SilentlyContinue
    Remove-Item Env:OBS_TLS_KEY -ErrorAction SilentlyContinue
    Remove-Item Env:OBS_ALLOW_INSECURE_HTTP -ErrorAction SilentlyContinue
}
