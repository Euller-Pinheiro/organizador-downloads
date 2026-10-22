<#
.SYNOPSIS
    Organiza a pasta Downloads em categorias, separa duplicatas e versões antigas — sem apagar nada.

.DESCRIPTION
    Organizar-Downloads lê os itens soltos na raiz de uma pasta (por padrão, a sua Downloads)
    e move cada um para uma categoria pelo tipo e pelo nome do arquivo.

    - Simulação por padrão: mostra e grava o plano; só move com -Executar.
    - Nunca apaga: duplicatas idênticas (mesmo SHA-256) e arquivos vazios vão para "_Revisar e apagar".
    - Versões: "relatorio-v2.pdf" e "relatorio-v3.pdf" -> fica a v3, a v2 vai para "Versões antigas".
    - Desfazer: cada execução grava um log CSV e um script que devolve tudo para o lugar.
    - Idempotente: rodar de novo só mexe no que chegou depois.
    - Funciona no Windows PowerShell 5.1 e no PowerShell 7 (Windows, Linux, macOS).

.PARAMETER Pasta
    Pasta a organizar. Padrão: Downloads do usuário.

.PARAMETER Executar
    Sem ele, nada é movido (simulação).

.PARAMETER SemVersoes
    Não separa versões antigas (mantém v1, v2, v3... lado a lado).

.EXAMPLE
    .\Organizar-Downloads.ps1                 # simulação
    .\Organizar-Downloads.ps1 -Executar       # organiza de verdade

.LINK
    https://github.com/Euller-Pinheiro/organizador-downloads
#>
[CmdletBinding()]
param(
    [string]$Pasta = (Join-Path $HOME 'Downloads'),
    [switch]$Executar,
    [switch]$SemVersoes
)

$ErrorActionPreference = 'Stop'
$Pasta = (Resolve-Path -LiteralPath $Pasta).Path
$PastaLog = Join-Path $Pasta '_Organização'
$Agora = Get-Date -Format 'yyyyMMdd-HHmmss'

# ============================================================================
#  REGRAS
#  Cada regra: expressão regular (sobre o nome do arquivo, sem acentos e sem
#  diferenciar maiúsculas) -> pasta de destino. A primeira que casar vence.
# ============================================================================
function Nova-Regra([string]$Re, [string]$Dest) { [pscustomobject]@{ Re = $Re; Dest = $Dest } }

# --- SUAS REGRAS: coloque aqui as específicas, elas vêm antes das genéricas.
$RegrasPessoais = @(
    # Nova-Regra 'cliente-acme|acme'          'Clientes\Acme'
    # Nova-Regra '^proposta'                  'Comercial\Propostas'
    # Nova-Regra 'nota.?fiscal|^nf[-_ ]'      'Financeiro\Notas fiscais'
)

# --- Regras genéricas, por tipo de arquivo
$RegrasGenericas = @(
    (Nova-Regra '^~\$'                                                       '_Revisar e apagar\Temporários do Office'),
    (Nova-Regra '\.(pfx|p12|pem|key|ovpn|kdbx|ppk)$'                         'Sensível - Certificados e Chaves'),
    (Nova-Regra '\.(exe|msi|msix|appx|iso|img|dmg|pkg|deb|rpm|apk)$'         'Instaladores'),
    (Nova-Regra '\.(zip|rar|7z|tar|gz|tgz|bz2|xz)$'                          'Compactados'),
    (Nova-Regra '\.pdf$'                                                     'Documentos\PDF'),
    (Nova-Regra '\.(docx?|odt|rtf|txt|md)$'                                  'Documentos\Texto'),
    (Nova-Regra '\.(xlsx?|ods|csv|tsv)$'                                     'Planilhas'),
    (Nova-Regra '\.(pptx?|odp|key)$'                                         'Apresentações'),
    (Nova-Regra '\.(png|jpe?g|gif|webp|svg|heic|bmp|tiff?|ico)$'             'Imagens'),
    (Nova-Regra '\.(mp4|mov|mkv|webm|avi|wmv)$'                              'Vídeos'),
    (Nova-Regra '\.(mp3|wav|ogg|m4a|flac|aac)$'                              'Áudio'),
    (Nova-Regra '\.(ttf|otf|woff2?)$'                                        'Fontes'),
    (Nova-Regra '\.(py|js|ts|php|ps1|psm1|sh|bat|cmd|sql|html?|css|json|ya?ml|xml|java|c|cpp|cs|go|rb|ipynb)$' 'Código')
)

$Regras = @($RegrasPessoais) + @($RegrasGenericas)

$REV = '_Revisar (sem categoria)'
$LIX = '_Revisar e apagar'
$ANT = 'Versões antigas'

# Pastas criadas por este script: nunca são movidas
$Raizes = @($Regras | ForEach-Object { ($_.Dest -split '\\')[0] }) + @($REV, $LIX, '_Organização', 'Pastas') | Select-Object -Unique

# ============================================================================
#  FUNÇÕES
# ============================================================================
function Sem-Acento([string]$t) {
    $d = $t.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $d.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($ch) }
    }
    $sb.ToString()
}

function Destino-Por-Regra([string]$nome) {
    $n = Sem-Acento $nome
    foreach ($r in $Regras) { if ($n -match $r.Re) { return $r.Dest } }
    return $REV
}

function Tem-MarcaDeCopia([string]$nome) {
    $b = [IO.Path]::GetFileNameWithoutExtension($nome)
    return ($b -match '\(\d+\)\s*$' -or $b -match '_\d$')
}

# "app-v12.3-final.zip" -> @{ Serie = "app--final.zip"; Versao = [version]12.3 }
function Ler-Versao([string]$nome) {
    $n = $nome.ToLowerInvariant()
    $m = [regex]::Match($n, '(?<![a-z0-9])v(\d+(?:\.\d+){0,3})(?![0-9])')
    if (-not $m.Success) { return $null }
    $txt = $m.Groups[1].Value
    if ($txt -notmatch '\.') { $txt = "$txt.0" }
    $serie = $n.Remove($m.Index, $m.Length)
    $serie = $serie -replace '\s*\(\d+\)', '' -replace '\d{4}-\d{2}-\d{2}', ''
    return [pscustomobject]@{ Serie = $serie; Versao = [version]$txt }
}

function Destino-Livre([string]$dir, [string]$nome) {
    $alvo = Join-Path $dir $nome
    if (-not (Test-Path -LiteralPath $alvo)) { return $alvo }
    $b = [IO.Path]::GetFileNameWithoutExtension($nome); $x = [IO.Path]::GetExtension($nome)
    for ($i = 2; ; $i++) {
        $alvo = Join-Path $dir ("{0} ({1}){2}" -f $b, $i, $x)
        if (-not (Test-Path -LiteralPath $alvo)) { return $alvo }
    }
}

function Caminho-Local([string]$rel) { Join-Path $Pasta ($rel -replace '\\', [IO.Path]::DirectorySeparatorChar) }

# ============================================================================
#  INVENTÁRIO
# ============================================================================
$euMesmo = $MyInvocation.MyCommand.Path
$itens = Get-ChildItem -LiteralPath $Pasta -Force | Where-Object {
    $_.Name -ne 'desktop.ini' -and $_.Name -ne '.DS_Store' -and $_.FullName -ne $euMesmo -and $Raizes -notcontains $_.Name
}
$arquivos = @($itens | Where-Object { -not $_.PSIsContainer })
$pastas   = @($itens | Where-Object { $_.PSIsContainer })

Write-Host ""
Write-Host "Pasta: $Pasta"
Write-Host ("Itens soltos na raiz: {0} arquivos, {1} pastas" -f $arquivos.Count, $pastas.Count)

$plano = New-Object System.Collections.ArrayList
$tratado = @{}
# Tamanho de uma pasta = soma dos arquivos dentro dela (para o relatório).
function Tamanho-Item($item) {
    if (-not $item.PSIsContainer) { return [int64]$item.Length }
    $soma = (Get-ChildItem -LiteralPath $item.FullName -Recurse -File -Force -ErrorAction SilentlyContinue |
        Measure-Object -Property Length -Sum).Sum
    if ($null -eq $soma) { return [int64]0 }
    return [int64]$soma
}

# Cada item do plano guarda POR QUE foi para onde vai (Tipo), além do destino.
# Tipo: Vazio | Duplicata | Versao | Categoria | SemCategoria | Pasta
function Planejar($item, [string]$dest, [string]$motivo, [string]$tipo) {
    [void]$plano.Add([pscustomobject]@{
        Origem     = $item.FullName
        Nome       = $item.Name
        DestDir    = $dest
        Motivo     = $motivo
        Tipo       = $tipo
        Tamanho    = (Tamanho-Item $item)
        Modificado = $item.LastWriteTime
    })
    $tratado[$item.FullName] = $true
}

# 1) Vazios
foreach ($f in $arquivos) { if ($f.Length -eq 0) { Planejar $f "$LIX\Arquivos vazios" 'arquivo vazio (0 bytes)' 'Vazio' } }

# 2) Duplicatas idênticas: mesmo tamanho e mesmo SHA-256
Write-Host "Procurando duplicatas idênticas..."
$nDup = 0
foreach ($g in ($arquivos | Where-Object { -not $tratado.ContainsKey($_.FullName) } | Group-Object Length | Where-Object { $_.Count -gt 1 })) {
    $porHash = @{}
    foreach ($f in $g.Group) {
        try { $h = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash } catch { continue }
        if (-not $porHash.ContainsKey($h)) { $porHash[$h] = New-Object System.Collections.ArrayList }
        [void]$porHash[$h].Add($f)
    }
    foreach ($lista in $porHash.Values) {
        if ($lista.Count -lt 2) { continue }
        # Fica: o que tem categoria; o de nome sem "(1)"; o de nome mais curto; o mais antigo
        $fica = $lista | Sort-Object @{E={ [int]((Destino-Por-Regra $_.Name) -eq $REV) }}, @{E={ [int](Tem-MarcaDeCopia $_.Name) }}, @{E={ $_.Name.Length }}, LastWriteTime | Select-Object -First 1
        foreach ($f in $lista) {
            if ($f.FullName -ne $fica.FullName) { Planejar $f "$LIX\Duplicatas idênticas" "cópia idêntica de: $($fica.Name)" 'Duplicata'; $nDup++ }
        }
    }
}

# 3) Classificação + versões
$restantes = @($arquivos | Where-Object { -not $tratado.ContainsKey($_.FullName) } | ForEach-Object {
    [pscustomobject]@{ Arq = $_; Dest = (Destino-Por-Regra $_.Name); Ver = (Ler-Versao $_.Name) }
})
$nAnt = 0
if (-not $SemVersoes) {
    $comVersao = $restantes | Where-Object { $_.Ver } | Group-Object { $_.Dest + '|' + $_.Ver.Serie } | Where-Object { $_.Count -gt 1 }
    foreach ($g in $comVersao) {
        $ord = @($g.Group | Sort-Object @{E={ $_.Ver.Versao }; Descending = $true}, @{E={ $_.Arq.LastWriteTime }; Descending = $true})
        for ($i = 1; $i -lt $ord.Count; $i++) {
            Planejar $ord[$i].Arq "$($ord[$i].Dest)\$ANT" "versão antiga (mais recente: $($ord[0].Arq.Name))" 'Versao'
            $nAnt++
        }
    }
}
foreach ($c in $restantes) {
    if (-not $tratado.ContainsKey($c.Arq.FullName)) {
        $tipo = if ($c.Dest -eq $REV) { 'SemCategoria' } else { 'Categoria' }
        Planejar $c.Arq $c.Dest 'classificado por tipo/nome' $tipo
    }
}

# 4) Pastas soltas (ex.: zips já extraídos)
foreach ($d in $pastas) { Planejar $d 'Pastas' 'pasta solta' 'Pasta' }

# ============================================================================
#  RESUMO
# ============================================================================
Write-Host ""
Write-Host "=========== PLANO ===========" -ForegroundColor Cyan
$plano | Group-Object DestDir | Sort-Object Name | ForEach-Object { Write-Host ("{0,5}  {1}" -f $_.Count, $_.Name) }
Write-Host "-----------------------------"
Write-Host ("{0,5}  itens ({1} duplicatas idênticas, {2} versões antigas)" -f $plano.Count, $nDup, $nAnt)
Write-Host ""

if ($plano.Count -eq 0) { Write-Host "Nada para organizar." -ForegroundColor Green; return }
if (-not (Test-Path -LiteralPath $PastaLog)) { New-Item -ItemType Directory -Path $PastaLog | Out-Null }

if (-not $Executar) {
    $arq = Join-Path $PastaLog "plano-simulacao-$Agora.csv"
    $plano | Select-Object @{N='Item';E={ Split-Path -Leaf $_.Origem }}, @{N='Destino';E={ $_.DestDir }}, Motivo |
        Sort-Object Destino, Item | Export-Csv -LiteralPath $arq -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Host "SIMULAÇÃO: nada foi movido." -ForegroundColor Yellow
    Write-Host "Plano detalhado: $arq"
    Write-Host "Para organizar de verdade, rode de novo com -Executar"
    return
}

# ============================================================================
#  EXECUÇÃO
# ============================================================================
$log = New-Object System.Collections.ArrayList
$ok = 0; $erros = 0
foreach ($p in $plano) {
    $destDir = Caminho-Local $p.DestDir
    try {
        if (-not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
        $alvo = Destino-Livre $destDir (Split-Path -Leaf $p.Origem)
        Move-Item -LiteralPath $p.Origem -Destination $alvo
        [void]$log.Add([pscustomobject]@{ Status = 'OK'; Origem = $p.Origem; Destino = $alvo; Motivo = $p.Motivo }); $ok++
    } catch {
        [void]$log.Add([pscustomobject]@{ Status = 'ERRO'; Origem = $p.Origem; Destino = $destDir; Motivo = $_.Exception.Message }); $erros++
    }
}
$arqLog = Join-Path $PastaLog "log-$Agora.csv"
$log | Export-Csv -LiteralPath $arqLog -NoTypeInformation -Encoding UTF8 -Delimiter ';'

$desfazer = @'
# Desfazer.ps1 — devolve para a raiz tudo o que o Organizar-Downloads moveu (usa o log mais recente)
param([string]$Log = "")
$aqui = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($Log)) {
    $Log = (Get-ChildItem -LiteralPath $aqui -Filter 'log-*.csv' | Sort-Object Name -Descending | Select-Object -First 1).FullName
}
Write-Host "Desfazendo a partir de: $Log"
$linhas = @(Import-Csv -LiteralPath $Log -Delimiter ';' | Where-Object { $_.Status -eq 'OK' })
[array]::Reverse($linhas)
$n = 0
foreach ($l in $linhas) {
    if ((Test-Path -LiteralPath $l.Destino) -and -not (Test-Path -LiteralPath $l.Origem)) {
        Move-Item -LiteralPath $l.Destino -Destination $l.Origem; $n++
    } else { Write-Host "Pulado: $($l.Destino)" -ForegroundColor Yellow }
}
Rename-Item -LiteralPath $Log -NewName ((Split-Path -Leaf $Log) -replace '^log-', 'desfeito-')
Write-Host "$n itens devolvidos."
'@
$arqDesfazer = Join-Path $PastaLog 'Desfazer.ps1'
[IO.File]::WriteAllText($arqDesfazer, $desfazer, (New-Object System.Text.UTF8Encoding($true)))

Write-Host ("Concluído: {0} itens movidos, {1} erros." -f $ok, $erros) -ForegroundColor Green
Write-Host "Log: $arqLog"
Write-Host "Para desfazer: `"$arqDesfazer`""
if ($erros -gt 0) { Write-Host "Itens com ERRO (geralmente arquivo aberto) ficaram na raiz; feche-os e rode de novo." -ForegroundColor Yellow }
