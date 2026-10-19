# Organizador de Downloads

[![Testes](https://github.com/Euller-Pinheiro/organizador-downloads/actions/workflows/testes.yml/badge.svg)](https://github.com/Euller-Pinheiro/organizador-downloads/actions/workflows/testes.yml)
![PowerShell 5.1 | 7](https://img.shields.io/badge/PowerShell-5.1%20%7C%207-5391FE?logo=powershell&logoColor=white)
![Licença MIT](https://img.shields.io/badge/licen%C3%A7a-MIT-green)

**[English below](#english)**

Script em PowerShell que organiza a pasta **Downloads** em categorias. Ele também separa as **duplicatas** e as **versões antigas**, e **nunca apaga nada**.

Nasceu de um problema real: eu tinha 858 arquivos soltos na minha pasta Downloads, em 20 GB, com dezenas de `arquivo (1).pdf` e pacotes `app-v9.zip` até `app-v55.zip`. Depois de rodar o script, sobraram só as pastas de categoria. Cada movimento ficou registrado em log, com um botão de desfazer.

## O que ele faz

```
Antes                              Depois
Downloads/                         Downloads/
├── contrato.pdf                   ├── Documentos/PDF/contrato.pdf
├── contrato (1).pdf   ← idêntico  ├── Compactados/app-v1.10.zip
├── app-v1.2.zip                   │   └── Versões antigas/app-v1.2.zip
├── app-v1.10.zip                  ├── Instaladores/setup-vscode.exe
├── setup-vscode.exe               ├── Sensível - Certificados e Chaves/certificado-a1.pfx
├── certificado-a1.pfx             ├── Planilhas/planilha-custos.xlsx
├── planilha-custos.xlsx           ├── _Revisar e apagar/
├── vazio.txt                      │   ├── Duplicatas idênticas/contrato (1).pdf
└── arquivo-estranho.xyz           │   └── Arquivos vazios/vazio.txt
                                   └── _Revisar (sem categoria)/arquivo-estranho.xyz
```

- **Categorias por tipo:** PDF, documentos de texto, planilhas, apresentações, imagens, vídeos, áudio, compactados, instaladores, código, fontes. Certificados e chaves (`.pfx`, `.pem`, `.ovpn`, `.kdbx`…) vão para uma pasta própria, marcada como **Sensível**.
- **Duplicatas de verdade:** compara o conteúdo pelo hash SHA-256, não pelo nome. Entre cópias idênticas, fica a que tem nome "limpo" e categoria. As outras vão para `_Revisar e apagar`.
- **Versões:** reconhece `v2`, `v2.1`, `v10` no nome e compara como número, então `v1.10` é mais nova que `v1.2`. Fica a mais recente, e as anteriores vão para `Versões antigas`.
- **Funciona com qualquer nome:** acentos (inclusive os que vêm do macOS), colchetes, espaços. Se já existir um arquivo com o mesmo nome no destino, o novo vira `nome (2).ext` e nada é sobrescrito.

## Garantias

| | |
|---|---|
| **Simulação por padrão** | Sem `-Executar`, nada é movido. Ele só mostra o plano e o grava em CSV, para você abrir no Excel. |
| **Nunca apaga** | Duplicatas e arquivos vazios vão para `_Revisar e apagar`, e quem decide apagar é você. |
| **Desfazer** | Toda execução grava `_Organização/log-*.csv` e um `Desfazer.ps1`, que devolve tudo para o lugar. |
| **Idempotente** | Rodar de novo só mexe no que chegou depois. |
| **Arquivo aberto** | É pulado, registrado como erro, e o resto continua normalmente. |

## Como usar

1. Baixe o [`Organizar-Downloads.ps1`](Organizar-Downloads.ps1) e salve na pasta Downloads.
2. Abra o PowerShell nessa pasta e rode a **simulação**:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Organizar-Downloads.ps1
   ```
3. Confira o plano. Se estiver como você quer, rode de verdade:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Organizar-Downloads.ps1 -Executar
   ```
4. Se não gostar do resultado:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\_Organização\Desfazer.ps1
   ```

| Parâmetro | O que faz |
|---|---|
| `-Pasta <caminho>` | Organiza outra pasta (o padrão é a Downloads do usuário). |
| `-Executar` | Move de verdade. Sem ele, só simula. |
| `-SemVersoes` | Não separa versões antigas. |

## Personalizando

No início do script existe um bloco `$RegrasPessoais`. As regras que você colocar ali têm prioridade sobre as genéricas:

```powershell
$RegrasPessoais = @(
    Nova-Regra 'cliente-acme|acme'        'Clientes\Acme'
    Nova-Regra '^proposta'                'Comercial\Propostas'
    Nova-Regra 'nota.?fiscal|^nf[-_ ]'    'Financeiro\Notas fiscais'
)
```

Cada regra é uma expressão regular aplicada ao nome do arquivo, sem acentos e sem diferenciar maiúsculas de minúsculas. A primeira que casar define o destino.

## Testes

[`tests/Testar.ps1`](tests/Testar.ps1) monta uma pasta de teste com os casos difíceis: duplicatas, versões `v2`/`v10`, nomes com acentos em NFD, colchetes, colisão de nomes, arquivo vazio e pasta solta. Depois ele verifica quatro coisas:

1. a simulação não move nada;
2. a execução põe cada arquivo no lugar certo **sem perder nenhum**, conferindo pelo hash;
3. rodar de novo não muda nada;
4. o desfazer devolve **exatamente** o estado original.

O GitHub Actions roda esses testes a cada push, no **Windows PowerShell 5.1**, no **PowerShell 7 (Windows)** e no **PowerShell 7 (Linux)**.

```powershell
pwsh ./tests/Testar.ps1
```

## Autor

**Euller Pinheiro**, Diretor Técnico da [ESX Tecnologia](https://esxtecnologia.com). A ESX cuida de TI gerenciada, infraestrutura e desenvolvimento de sistemas em São Paulo.

Precisa de uma automação parecida para a sua empresa? [Fale com a ESX](https://esxtecnologia.com).

---

<a name="english"></a>
## English

A PowerShell script that tidies up your **Downloads** folder: it sorts files into categories and separates **exact duplicates** (by SHA-256) and **older versions** (`v1.2` < `v1.10`, compared numerically). It **never deletes anything**.

- **Dry-run by default:** nothing moves without `-Executar`. The plan is saved as a CSV.
- **Never deletes:** duplicates and empty files go to `_Revisar e apagar` ("review and delete").
- **Undo:** every run writes a CSV log and a `Desfazer.ps1` ("undo") script.
- **Idempotent and safe with any filename:** accents (NFC/NFD), brackets, name collisions.
- Tested in CI on Windows PowerShell 5.1, PowerShell 7 on Windows, and PowerShell 7 on Linux.

```powershell
.\Organizar-Downloads.ps1             # dry-run
.\Organizar-Downloads.ps1 -Executar   # do it
```

Folder and message names are in Portuguese. You can add your own rules in `$RegrasPessoais`.

Licensed under [MIT](LICENSE).
