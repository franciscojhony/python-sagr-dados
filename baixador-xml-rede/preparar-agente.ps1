# Prepara o Python PORTÁTIL do Agente (rode UMA vez no servidor, com internet).
# Baixa o Python "embutido" (sem instalar nada), instala o selenium nele e deixa tudo em
#   agente-portatil\python
# O servidor passa a incluir essa pasta no pacote que os usuários baixam em /agente.zip.
$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ProgressPreference = "SilentlyContinue"

$ver  = "3.12.10"
$raiz = Join-Path $PSScriptRoot "agente-portatil"
$py   = Join-Path $raiz "python"
$zip  = Join-Path $env:TEMP "python-$ver-embed-amd64.zip"

if (Test-Path $py) { Remove-Item $py -Recurse -Force }
New-Item -ItemType Directory -Force -Path $py | Out-Null

Write-Host "1/4 Baixando o Python $ver (versao embutida)..."
Invoke-WebRequest "https://www.python.org/ftp/python/$ver/python-$ver-embed-amd64.zip" -OutFile $zip
Expand-Archive $zip -DestinationPath $py -Force

Write-Host "2/4 Habilitando pacotes (site-packages)..."
$pth = Get-ChildItem $py -Filter "python*._pth" | Select-Object -First 1
$zipPy = (Get-ChildItem $py -Filter "python*.zip" | Select-Object -First 1).Name
Set-Content -Path $pth.FullName -Value @($zipPy, ".", "Lib\site-packages", "import site") -Encoding ASCII

Write-Host "3/4 Instalando o pip..."
$gp = Join-Path $py "get-pip.py"
Invoke-WebRequest "https://bootstrap.pypa.io/get-pip.py" -OutFile $gp
& (Join-Path $py "python.exe") $gp --no-warn-script-location
Remove-Item $gp -Force

Write-Host "4/4 Instalando o selenium..."
& (Join-Path $py "python.exe") -m pip install selenium --no-warn-script-location

Write-Host ""
Write-Host "Pronto! O pacote portatil esta em: $py"
Write-Host "Agora os usuarios podem baixar o Agente em  http://IP-DO-SERVIDOR:8765/agente.zip"
