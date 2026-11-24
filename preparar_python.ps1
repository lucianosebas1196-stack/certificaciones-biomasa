# Prepara el Python portátil (se corre solo, una vez, desde _python.bat).
# Baja de python.org el Python oficial "embeddable" 3.13 de 64 bits (unos 11 MB) y lo pone en la carpeta python\.
# Las librerías del programa ya vienen en python\Lib\site-packages: no se instala nada en Windows.
# Sin internet: bajá a mano "Windows embeddable package (64-bit)" de Python 3.13 desde python.org/downloads/windows,
# dejalo en esta carpeta con el nombre python-embed.zip y volvé a abrir el programa.
$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$aqui = Split-Path -Parent $MyInvocation.MyCommand.Path
$destino = Join-Path $aqui "python"
$zip = Join-Path $aqui "python-embed.zip"
Write-Host ""
Write-Host " Preparando Python portatil (solo la primera vez)..."
if (-not (Test-Path $zip)) {
    $bajado = $false
    for ($n = 30; $n -ge 5; $n--) {
        $url = "https://www.python.org/ftp/python/3.13.$n/python-3.13.$n-embed-amd64.zip"
        try {
            Invoke-WebRequest -Uri $url -OutFile $zip -UseBasicParsing
            Write-Host " Bajado Python 3.13.$n"
            $bajado = $true
            break
        } catch {
            if (Test-Path $zip) { Remove-Item $zip -Force }
        }
    }
    if (-not $bajado) {
        Write-Host ""
        Write-Host " ERROR: no se pudo bajar Python de python.org (sin internet o bloqueado)."
        Write-Host " Baja a mano 'Windows embeddable package (64-bit)' de Python 3.13 desde"
        Write-Host " https://www.python.org/downloads/windows/ , guardalo en esta carpeta como python-embed.zip"
        Write-Host " y volve a abrir el programa."
        exit 1
    }
}
try {
    Expand-Archive -Path $zip -DestinationPath $destino -Force
} catch {
    Write-Host " ERROR: el archivo python-embed.zip esta danado. Borralo y volve a abrir el programa."
    exit 1
}
$pth = Get-ChildItem -Path $destino -Filter "python3*._pth" | Select-Object -First 1
if (-not $pth) {
    Write-Host " ERROR: python-embed.zip no es el Python 'embeddable' de Windows. Borralo y volve a abrir el programa."
    Remove-Item (Join-Path $destino "python.exe") -ErrorAction SilentlyContinue
    exit 1
}
if ($pth.Name -ne "python313._pth") {
    Write-Host " ERROR: el Python bajado no es la version 3.13 (es $($pth.Name)). Las librerias son para 3.13."
    Write-Host " Borra python-embed.zip y bajate el de Python 3.13."
    Remove-Item (Join-Path $destino "python.exe") -ErrorAction SilentlyContinue
    exit 1
}
$base = [IO.Path]::GetFileNameWithoutExtension($pth.Name)
Set-Content -Path $pth.FullName -Encoding ASCII -Value @("$base.zip", ".", "Lib\site-packages", "import site")
Remove-Item $zip -Force
Write-Host " Listo: Python portatil preparado."
Write-Host ""
