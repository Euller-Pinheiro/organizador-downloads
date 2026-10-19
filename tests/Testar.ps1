# Teste de ponta a ponta do Organizar-Downloads.ps1 (sem dependências).
# Uso: pwsh ./tests/Testar.ps1     (ou: powershell -File .\tests\Testar.ps1)
# Sai com código 1 se qualquer verificação falhar.

$ErrorActionPreference = 'Stop'
$raiz    = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$script  = Join-Path $raiz 'Organizar-Downloads.ps1'
$base    = Join-Path ([IO.Path]::GetTempPath()) ("org-teste-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $base | Out-Null
$falhas  = 0
$sep     = [IO.Path]::DirectorySeparatorChar

function Ok([bool]$cond, [string]$msg) {
    if ($cond) { Write-Host "  [ok]    $msg" -ForegroundColor Green }
    else       { Write-Host "  [FALHA] $msg" -ForegroundColor Red; $script:falhas++ }
}
function Escrever([string]$nome, [string]$conteudo) {
    $p = Join-Path $base $nome
    [IO.File]::WriteAllText($p, $conteudo)
    return $p
}
function Rel([string]$r) { Join-Path $base ($r -replace '\\', $sep) }
function Retrato {
    Get-ChildItem -LiteralPath $base -Recurse -File -Force |
        Where-Object { $_.FullName -notlike "*_Organização*" } |
        ForEach-Object { "{0}|{1}" -f $_.FullName.Substring($base.Length), (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash } |
        Sort-Object
}
function Hashes {
    Get-ChildItem -LiteralPath $base -Recurse -File -Force |
        Where-Object { $_.FullName -notlike "*_Organização*" } |
        ForEach-Object { (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash } | Sort-Object
}
function Rodar([switch]$Executar) {
    if ($Executar) { & $script -Pasta $base -Executar *> $null } else { & $script -Pasta $base *> $null }
}

# ---------------------------------------------------------------- cenário
$nfd = "Relato" + [char]0x0301 + "rio"           # "Relatório" decomposto (NFD), como vem de macOS
Escrever 'contrato.pdf'            'CONTRATO'        | Out-Null
Escrever 'contrato (1).pdf'        'CONTRATO'        | Out-Null   # duplicata idêntica
Escrever 'contrato (2).pdf'        'CONTRATO v2 dif' | Out-Null   # mesmo nome-base, conteúdo diferente -> fica
Escrever 'app-v1.zip'              'z1'              | Out-Null
Escrever 'app-v2.zip'              'z2'              | Out-Null
Escrever 'app-v10.zip'             'z10'             | Out-Null   # v10 > v2 (comparação numérica)
Escrever 'proposta-v2.1.docx'      'p21'             | Out-Null
Escrever 'proposta-v2.docx'        'p2'              | Out-Null
Escrever 'certificado.pfx'         'CERT'            | Out-Null
Escrever 'setup.exe'               'EXE'             | Out-Null
Escrever "$nfd anual.xlsx"         'XLSX'            | Out-Null
Escrever 'foto.jpg'                'FOTO-NOVA'       | Out-Null
Escrever 'desconhecido.xyz'        'XYZ'             | Out-Null
Escrever '[colchetes] nota.pdf'    'COLCHETES'       | Out-Null
Escrever '~$trava.docx'            'TRAVA'           | Out-Null
New-Item -ItemType File -Path (Join-Path $base 'vazio.txt') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $base 'zip-extraido') | Out-Null
Escrever "zip-extraido${sep}dentro.txt" 'DENTRO' | Out-Null
# categoria já existente com um arquivo de mesmo nome -> deve virar "foto (2).jpg"
New-Item -ItemType Directory -Path (Rel 'Imagens') | Out-Null
Escrever "Imagens${sep}foto.jpg" 'FOTO-ANTIGA' | Out-Null

$antes = Retrato
$hashesAntes = Hashes

# ---------------------------------------------------------------- 1. simulação
Write-Host "`n1. Simulação"
Rodar
Ok ((Compare-Object $antes (Retrato)) -eq $null) 'simulação não move nada'
Ok (@(Get-ChildItem -LiteralPath (Join-Path $base '_Organização') -Filter 'plano-simulacao-*.csv').Count -eq 1) 'plano CSV gravado'

# ---------------------------------------------------------------- 2. execução
Write-Host "`n2. Execução"
Rodar -Executar
$esperado = @{
    'Documentos\PDF\contrato.pdf'                           = $true
    'Documentos\PDF\contrato (2).pdf'                       = $true
    '_Revisar e apagar\Duplicatas idênticas\contrato (1).pdf' = $true
    'Compactados\app-v10.zip'                               = $true
    'Compactados\Versões antigas\app-v2.zip'                = $true
    'Compactados\Versões antigas\app-v1.zip'                = $true
    'Documentos\Texto\proposta-v2.1.docx'                   = $true
    'Documentos\Texto\Versões antigas\proposta-v2.docx'     = $true
    'Sensível - Certificados e Chaves\certificado.pfx'      = $true
    'Instaladores\setup.exe'                                = $true
    "Planilhas\$nfd anual.xlsx"                             = $true
    'Imagens\foto (2).jpg'                                  = $true
    'Imagens\foto.jpg'                                      = $true
    '_Revisar (sem categoria)\desconhecido.xyz'             = $true
    'Documentos\PDF\[colchetes] nota.pdf'                   = $true
    '_Revisar e apagar\Temporários do Office\~$trava.docx'  = $true
    '_Revisar e apagar\Arquivos vazios\vazio.txt'           = $true
    'Pastas\zip-extraido\dentro.txt'                        = $true
}
foreach ($k in $esperado.Keys) { Ok (Test-Path -LiteralPath (Rel $k)) "existe: $k" }
Ok ((Get-Content -LiteralPath (Rel 'Imagens\foto.jpg') -Raw) -eq 'FOTO-ANTIGA') 'arquivo pré-existente na categoria não foi sobrescrito'
$soltos = @(Get-ChildItem -LiteralPath $base -File -Force)
Ok ($soltos.Count -eq 0) "raiz sem arquivos soltos (restaram: $($soltos.Name -join ', '))"
Ok ((Compare-Object $hashesAntes (Hashes)) -eq $null) 'nenhum arquivo perdido (mesmos conteúdos, mesma quantidade)'

# ---------------------------------------------------------------- 3. idempotência
Write-Host "`n3. Rodar de novo"
$depois = Retrato
Rodar -Executar
Ok ((Compare-Object $depois (Retrato)) -eq $null) 'segunda execução não mexe em nada'

# ---------------------------------------------------------------- 4. desfazer
Write-Host "`n4. Desfazer"
& (Join-Path $base '_Organização\Desfazer.ps1'.Replace('\', $sep)) *> $null
Get-ChildItem -LiteralPath $base -Directory | Where-Object { $_.Name -ne '_Organização' -and $_.Name -ne 'zip-extraido' -and $_.Name -ne 'Imagens' } |
    Where-Object { -not (Get-ChildItem -LiteralPath $_.FullName -Recurse -File) } | Remove-Item -Recurse -Force
Ok ((Compare-Object $antes (Retrato)) -eq $null) 'desfazer devolve exatamente o estado original'

# ---------------------------------------------------------------- fim
Remove-Item -LiteralPath $base -Recurse -Force
Write-Host ""
if ($falhas -gt 0) { Write-Host "$falhas verificação(ões) falharam." -ForegroundColor Red; exit 1 }
Write-Host "Todos os testes passaram." -ForegroundColor Green
exit 0
