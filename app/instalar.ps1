# Instalador do Gerador de Propostas - Solar In Focco
$ErrorActionPreference = "Continue"
$Host.UI.RawUI.WindowTitle = "Instalando o Gerador de Propostas - Solar In Focco"
$raiz = Split-Path -Parent $MyInvocation.MyCommand.Path
$app  = $raiz
try { Start-Transcript -Path (Join-Path $app "instalacao.log") -Force | Out-Null } catch {}
trap { Write-Host ""; Write-Host "ERRO: $_" -ForegroundColor Red; Write-Host "Tire um print desta janela e envie."; Read-Host "Pressione ENTER para sair"; exit 1 }

function Passo($n, $txt) { Write-Host ""; Write-Host "[$n/5] $txt" -ForegroundColor Cyan }
function Ok($txt)  { Write-Host "      OK - $txt" -ForegroundColor Green }
function Aviso($txt) { Write-Host "      ATENCAO - $txt" -ForegroundColor Yellow }

Write-Host "==============================================================" -ForegroundColor DarkCyan
Write-Host "   Gerador de Propostas - Solar In Focco  |  Instalacao"       -ForegroundColor White
Write-Host "==============================================================" -ForegroundColor DarkCyan
Write-Host "Isso pode levar alguns minutos. Se o Windows perguntar se pode"
Write-Host "fazer alteracoes no computador, clique em SIM."

$temWinget = [bool](Get-Command winget -ErrorAction SilentlyContinue)
if (-not $temWinget) {
    Aviso "O 'Instalador de Aplicativo' do Windows nao foi encontrado."
    Write-Host "      Abra a Microsoft Store, procure 'Instalador de Aplicativo', clique em Atualizar e rode este instalador de novo."
    Start-Process "ms-windows-store://pdp/?productid=9NBLGGH4NNS1"
    Read-Host "Pressione ENTER para sair"; exit 1
}
$wingetArgs = @("--accept-package-agreements", "--accept-source-agreements", "--silent", "-e")

# ---------- 1. Python ----------
function Achar-Python {
    $c = @(
        "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python313\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe",
        "$env:ProgramFiles\Python312\python.exe",
        "$env:ProgramFiles\Python313\python.exe",
        "$env:ProgramFiles\Python311\python.exe"
    )
    foreach ($p in $c) { if (Test-Path $p) { return $p } }
    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) {
        $exe = & py -3 -c "import sys;print(sys.executable)" 2>$null
        if ($exe -and (Test-Path $exe) -and ($exe -notlike "*WindowsApps*")) { return $exe }
    }
    return $null
}
Passo 1 "Python"
$python = Achar-Python
if (-not $python) {
    Write-Host "      Instalando o Python..."
    winget install --id Python.Python.3.12 --scope user @wingetArgs | Out-Null
    $python = Achar-Python
}
if (-not $python) { Aviso "Nao consegui instalar o Python. Instale pelo site python.org e rode de novo."; Read-Host "ENTER para sair"; exit 1 }
Ok $python

# ---------- 2. Bibliotecas ----------
Passo 2 "Componentes do programa"
& $python -m pip install --upgrade pip --quiet --disable-pip-version-check 2>$null | Out-Null
& $python -m pip install openpyxl pdfplumber pillow lxml pypdfium2 --quiet --disable-pip-version-check
if ($LASTEXITCODE -ne 0) { Aviso "Falha ao instalar componentes. Verifique a internet e rode de novo."; Read-Host "ENTER para sair"; exit 1 }
Ok "componentes instalados"

# ---------- 3. LibreOffice ----------
Passo 3 "LibreOffice (calcula a planilha e gera o PDF)"
$soffice = @("$env:ProgramFiles\LibreOffice\program\soffice.exe", "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $soffice) {
    Write-Host "      Instalando o LibreOffice (download grande, aguarde)..."
    winget install --id TheDocumentFoundation.LibreOffice @wingetArgs | Out-Null
    $soffice = @("$env:ProgramFiles\LibreOffice\program\soffice.exe", "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
}
if ($soffice) { Ok "LibreOffice pronto" } else { Aviso "LibreOffice nao foi instalado. Instale pelo site libreoffice.org." }

# ---------- 4. Tesseract (leitura de orcamento escaneado) ----------
Passo 4 "Leitor de documentos escaneados (Tesseract)"
$tess = @("$env:ProgramFiles\Tesseract-OCR\tesseract.exe", "$env:LOCALAPPDATA\Programs\Tesseract-OCR\tesseract.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $tess) {
    Write-Host "      Instalando o Tesseract..."
    winget install --id UB-Mannheim.TesseractOCR @wingetArgs | Out-Null
    $tess = @("$env:ProgramFiles\Tesseract-OCR\tesseract.exe", "$env:LOCALAPPDATA\Programs\Tesseract-OCR\tesseract.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
}
if ($tess) { Ok "leitor pronto (idioma portugues incluso)" } else { Aviso "Tesseract nao foi instalado. Orcamentos escaneados nao serao lidos." }

# ---------- 5. Atalho na Area de Trabalho ----------
Passo 5 "Criando o atalho na Area de Trabalho"
$pythonw = Join-Path (Split-Path $python) "pythonw.exe"
if (-not (Test-Path $pythonw)) { $pythonw = $python }
[IO.File]::WriteAllText((Join-Path $app "python_path.txt"), $pythonw)
$desk = [Environment]::GetFolderPath("Desktop")
$ws = New-Object -ComObject WScript.Shell
$lnk = $ws.CreateShortcut((Join-Path $desk "Gerar Proposta Solar In Focco.lnk"))
$lnk.TargetPath = $pythonw
$lnk.Arguments = '"' + (Join-Path $app "gerador_app.pyw") + '"'
$lnk.WorkingDirectory = $app
$lnk.IconLocation = (Join-Path $app "icone.ico")
$lnk.Description = "Gerador de Propostas - Solar In Focco"
$lnk.Save()
Ok "atalho 'Gerar Proposta Solar In Focco' criado"

Write-Host ""
Write-Host "==============================================================" -ForegroundColor Green
Write-Host "  PRONTO! Use o icone 'Gerar Proposta Solar In Focco'"          -ForegroundColor Green
Write-Host "  na sua Area de Trabalho."                                       -ForegroundColor Green
Write-Host "==============================================================" -ForegroundColor Green
$abrir = Read-Host "Deseja abrir o programa agora? (S/N)"
if ($abrir -match "^[sS]") { Start-Process $pythonw -ArgumentList ('"' + (Join-Path $app "gerador_app.pyw") + '"') -WorkingDirectory $app }
