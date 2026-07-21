# ============================================================
#  INVENTARIO DE MAQUINAS - Coletor (Windows)
#  Execute clicando duas vezes em "coletor.bat"
# ============================================================

# --------- CONFIGURACAO ---------
# Cole aqui a URL do Web App do Google Apps Script (ver apps-script-inventario.gs).
# Deixe como esta se quiser usar so o relatorio local, sem enviar para a planilha.
$APPS_SCRIPT_URL = 'COLE_AQUI_A_URL_DO_SEU_APPS_SCRIPT'

$ErrorActionPreference = 'SilentlyContinue'
$ScriptVersao = '1.0'

function Escreve-Etapa($texto) {
    Write-Host ">> $texto" -ForegroundColor Cyan
}

function Texto($valor) {
    if ($null -eq $valor) { return '' }
    return (($valor.ToString()) -replace '\s+', ' ').Trim()
}

Clear-Host
Write-Host "============================================================" -ForegroundColor DarkCyan
Write-Host "   INVENTARIO DE MAQUINAS - Coletando informacoes..." -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor DarkCyan
Write-Host ""

# --------- IDENTIFICACAO / SISTEMA ---------
Escreve-Etapa "Identificacao do computador..."
$cs  = Get-CimInstance Win32_ComputerSystem
$bios = Get-CimInstance Win32_BIOS
$enclosure = Get-CimInstance Win32_SystemEnclosure
$os  = Get-CimInstance Win32_OperatingSystem

$dominio = if ($cs.PartOfDomain) { "Dominio: $($cs.Domain)" } else { "Grupo de Trabalho: $($cs.Domain)" }

$patrimonio = $enclosure.SMBIOSAssetTag
if ([string]::IsNullOrWhiteSpace($patrimonio) -or $patrimonio -match '(?i)no asset|not specified|default string') {
    $patrimonio = ''
}

# --------- CPU ---------
Escreve-Etapa "Processador..."
$cpuInfo = Get-CimInstance Win32_Processor | Select-Object -First 1
$cpu = [ordered]@{
    nome      = Texto $cpuInfo.Name
    nucleos   = $cpuInfo.NumberOfCores
    threads   = $cpuInfo.NumberOfLogicalProcessors
    clockMHz  = $cpuInfo.MaxClockSpeed
}

# --------- MEMORIA RAM ---------
Escreve-Etapa "Memoria RAM..."
$ramModulosRaw = Get-CimInstance Win32_PhysicalMemory
$ramModulos = @()
foreach ($m in $ramModulosRaw) {
    $ramModulos += [ordered]@{
        capacidadeGB = [math]::Round($m.Capacity / 1GB, 1)
        velocidadeMHz = $m.Speed
        fabricante    = Texto $m.Manufacturer
        modelo        = Texto $m.PartNumber
    }
}
$ramTotalGB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)

# --------- DISCOS ---------
Escreve-Etapa "Discos e armazenamento..."
$discosFisicos = Get-CimInstance Win32_DiskDrive
$tiposDisco = @{}
try {
    Get-PhysicalDisk | ForEach-Object { $tiposDisco[$_.DeviceId] = $_.MediaType }
} catch {}

$discos = @()
foreach ($d in $discosFisicos) {
    $tipo = 'N/A'
    if ($d.Index -ne $null -and $tiposDisco.ContainsKey([string]$d.Index)) {
        $tipo = $tiposDisco[[string]$d.Index]
    }
    $discos += [ordered]@{
        modelo        = Texto $d.Model
        tamanhoGB     = if ($d.Size) { [math]::Round($d.Size / 1GB, 1) } else { 0 }
        tipo          = $tipo
        interface     = $d.InterfaceType
        numeroSerie   = Texto $d.SerialNumber
    }
}

$volumes = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3"
$espacoLivre = @()
foreach ($v in $volumes) {
    $espacoLivre += [ordered]@{
        unidade      = $v.DeviceID
        tamanhoGB    = [math]::Round($v.Size / 1GB, 1)
        livreGB      = [math]::Round($v.FreeSpace / 1GB, 1)
    }
}

# --------- PLACA DE VIDEO ---------
Escreve-Etapa "Placa de video..."
$gpusRaw = Get-CimInstance Win32_VideoController
$gpus = @()
foreach ($g in $gpusRaw) {
    $gpus += [ordered]@{
        nome           = Texto $g.Name
        driverVersao   = $g.DriverVersion
        vramMB         = if ($g.AdapterRAM) { [math]::Round($g.AdapterRAM / 1MB) } else { $null }
    }
}

# --------- REDE ---------
Escreve-Etapa "Configuracao de rede..."
$adaptadores = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True"
$rede = @()
foreach ($a in $adaptadores) {
    $rede += [ordered]@{
        nome     = $a.Description
        ip       = if ($a.IPAddress) { $a.IPAddress[0] } else { '' }
        mac      = $a.MACAddress
        gateway  = if ($a.DefaultIPGateway) { $a.DefaultIPGateway[0] } else { '' }
    }
}

# --------- SOFTWARES INSTALADOS ---------
Escreve-Etapa "Softwares instalados (pode levar alguns segundos)..."
$caminhosRegistro = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
)
$softwares = @()
foreach ($caminho in $caminhosRegistro) {
    Get-ItemProperty -Path $caminho -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.DisplayName -and $_.SystemComponent -ne 1 -and -not $_.ParentKeyName) {
            $softwares += [ordered]@{
                nome        = $_.DisplayName.Trim()
                versao      = $_.DisplayVersion
                fabricante  = $_.Publisher
                instalado   = $_.InstallDate
            }
        }
    }
}
$softwares = $softwares | Sort-Object -Property nome -Unique

# --------- MONTA OBJETO FINAL ---------
$dados = [ordered]@{
    timestamp           = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    hostname            = $env:COMPUTERNAME
    usuarioLogado       = $env:USERNAME
    dominio             = $dominio
    fabricante          = Texto $cs.Manufacturer
    modelo              = Texto $cs.Model
    numeroSerie         = Texto $bios.SerialNumber
    patrimonio          = $patrimonio
    cpu                 = $cpu
    ramTotalGB          = $ramTotalGB
    ramModulos          = $ramModulos
    discos              = $discos
    espacoLivre         = $espacoLivre
    gpus                = $gpus
    sistemaOperacional  = [ordered]@{
        nome         = Texto $os.Caption
        versao       = $os.Version
        build        = $os.BuildNumber
        arquitetura  = $os.OSArchitecture
        instalado    = if ($os.InstallDate) { $os.InstallDate.ToString('yyyy-MM-dd') } else { '' }
    }
    rede                = $rede
    softwares           = $softwares
    scriptVersao        = $ScriptVersao
}

# --------- SALVA RELATORIO LOCAL ---------
Escreve-Etapa "Gerando relatorio local..."
$pastaSaida = Join-Path $env:USERPROFILE 'Desktop\Inventario'
if (-not (Test-Path $pastaSaida)) { New-Item -ItemType Directory -Path $pastaSaida | Out-Null }

$carimbo = (Get-Date).ToString('yyyyMMdd_HHmmss')
$nomeBase = "$($dados.hostname)_$carimbo"
$caminhoJson = Join-Path $pastaSaida "$nomeBase.json"
$caminhoHtml = Join-Path $pastaSaida "$nomeBase.html"

$json = $dados | ConvertTo-Json -Depth 8
$json | Out-File -FilePath $caminhoJson -Encoding UTF8

function Escapa($texto) {
    if ($null -eq $texto) { return '' }
    return [System.Web.HttpUtility]::HtmlEncode($texto.ToString())
}
Add-Type -AssemblyName System.Web

$linhasRam = ($dados.ramModulos | ForEach-Object { "<tr><td>$(Escapa $_.fabricante) $(Escapa $_.modelo)</td><td>$($_.capacidadeGB) GB</td><td>$($_.velocidadeMHz) MHz</td></tr>" }) -join "`n"
$linhasDiscos = ($dados.discos | ForEach-Object { "<tr><td>$(Escapa $_.modelo)</td><td>$($_.tamanhoGB) GB</td><td>$(Escapa $_.tipo)</td><td>$(Escapa $_.interface)</td></tr>" }) -join "`n"
$linhasGpu = ($dados.gpus | ForEach-Object { "<tr><td>$(Escapa $_.nome)</td><td>$(Escapa $_.driverVersao)</td><td>$($_.vramMB) MB</td></tr>" }) -join "`n"
$linhasRede = ($dados.rede | ForEach-Object { "<tr><td>$(Escapa $_.nome)</td><td>$(Escapa $_.ip)</td><td>$(Escapa $_.mac)</td><td>$(Escapa $_.gateway)</td></tr>" }) -join "`n"
$linhasSoftware = ($dados.softwares | ForEach-Object { "<tr><td>$(Escapa $_.nome)</td><td>$(Escapa $_.versao)</td><td>$(Escapa $_.fabricante)</td></tr>" }) -join "`n"

$html = @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Inventario - $(Escapa $dados.hostname)</title>
<style>
  * { box-sizing:border-box; margin:0; padding:0; }
  body { font-family:'Segoe UI',Arial,sans-serif; background:#f1f3f8; color:#1a1a2e; padding:2rem; }
  .container { max-width:900px; margin:0 auto; }
  h1 { font-size:1.6rem; color:#0f2557; margin-bottom:.2rem; }
  .sub { color:#6b7280; font-size:.85rem; margin-bottom:1.5rem; }
  .grid { display:grid; grid-template-columns:repeat(2,1fr); gap:10px; margin-bottom:1.5rem; }
  .box { background:#fff; border:1px solid #dfe3ee; border-radius:10px; padding:.9rem 1.1rem; }
  .box .lbl { font-size:.72rem; text-transform:uppercase; letter-spacing:.5px; color:#6b7280; }
  .box .val { font-size:1rem; font-weight:600; color:#0f2557; margin-top:2px; }
  section { background:#fff; border:1px solid #dfe3ee; border-radius:10px; padding:1.1rem 1.3rem; margin-bottom:1rem; }
  section h2 { font-size:1rem; color:#0f2557; margin-bottom:.7rem; border-bottom:2px solid #eef1f8; padding-bottom:.4rem; }
  table { width:100%; border-collapse:collapse; font-size:.85rem; }
  th, td { text-align:left; padding:6px 8px; border-bottom:1px solid #eef1f8; }
  th { color:#6b7280; font-weight:600; font-size:.75rem; text-transform:uppercase; }
  footer { text-align:center; color:#9aa1b3; font-size:.75rem; margin-top:1.5rem; }
</style>
</head>
<body>
<div class="container">
  <h1>$(Escapa $dados.hostname)</h1>
  <div class="sub">Relatorio gerado em $($dados.timestamp) &middot; $(Escapa $dados.dominio) &middot; Usuario: $(Escapa $dados.usuarioLogado)</div>

  <div class="grid">
    <div class="box"><div class="lbl">Fabricante / Modelo</div><div class="val">$(Escapa $dados.fabricante) $(Escapa $dados.modelo)</div></div>
    <div class="box"><div class="lbl">Numero de Serie</div><div class="val">$(Escapa $dados.numeroSerie)</div></div>
    <div class="box"><div class="lbl">Processador</div><div class="val">$(Escapa $dados.cpu.nome)</div></div>
    <div class="box"><div class="lbl">Memoria RAM</div><div class="val">$($dados.ramTotalGB) GB</div></div>
    <div class="box"><div class="lbl">Sistema Operacional</div><div class="val">$(Escapa $dados.sistemaOperacional.nome)</div></div>
    <div class="box"><div class="lbl">Patrimonio</div><div class="val">$(Escapa $dados.patrimonio)</div></div>
  </div>

  <section><h2>Memoria RAM ($($dados.ramModulos.Count) modulo(s))</h2>
    <table><tr><th>Modulo</th><th>Capacidade</th><th>Velocidade</th></tr>$linhasRam</table></section>

  <section><h2>Discos</h2>
    <table><tr><th>Modelo</th><th>Tamanho</th><th>Tipo</th><th>Interface</th></tr>$linhasDiscos</table></section>

  <section><h2>Placa de Video</h2>
    <table><tr><th>Modelo</th><th>Driver</th><th>VRAM</th></tr>$linhasGpu</table></section>

  <section><h2>Rede</h2>
    <table><tr><th>Adaptador</th><th>IP</th><th>MAC</th><th>Gateway</th></tr>$linhasRede</table></section>

  <section><h2>Softwares Instalados ($($dados.softwares.Count))</h2>
    <table><tr><th>Nome</th><th>Versao</th><th>Fabricante</th></tr>$linhasSoftware</table></section>

  <footer>Inventario de Maquinas &middot; gerado automaticamente</footer>
</div>
</body>
</html>
"@
$html | Out-File -FilePath $caminhoHtml -Encoding UTF8

Write-Host ""
Write-Host "Relatorio salvo em:" -ForegroundColor Green
Write-Host "  $caminhoHtml"
Write-Host "  $caminhoJson"

# --------- ENVIA PARA A PLANILHA CENTRAL ---------
if ($APPS_SCRIPT_URL -and $APPS_SCRIPT_URL -notmatch 'COLE_AQUI') {
    Write-Host ""
    Escreve-Etapa "Enviando dados para a planilha central..."
    try {
        $resposta = Invoke-RestMethod -Uri $APPS_SCRIPT_URL -Method Post -Body $json -ContentType 'application/json; charset=utf-8' -TimeoutSec 30
        if ($resposta.sucesso) {
            Write-Host "Enviado com sucesso para o inventario central." -ForegroundColor Green
        } else {
            Write-Host "A planilha respondeu com erro: $($resposta.erro)" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Nao foi possivel enviar para a planilha central (sem internet ou URL invalida)." -ForegroundColor Yellow
        Write-Host "O relatorio local foi salvo normalmente." -ForegroundColor Yellow
    }
} else {
    Write-Host ""
    Write-Host "Envio automatico desativado (URL do Apps Script nao configurada)." -ForegroundColor DarkYellow
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor DarkCyan
Write-Host "   Concluido! Abrindo o relatorio..." -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor DarkCyan
Start-Process $caminhoHtml
