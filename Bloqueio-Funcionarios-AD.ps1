<#
    Bloqueio-Funcionarios-AD.ps1

    Le a planilha .xlsx e, numa unica execucao, desabilita no Active Directory:
      - Quem foi desligado de vez (RESCISAO)
      - Quem saiu de ferias, temporariamente (FERIAS) - ver Reativar-Ferias-AD.ps1

    ANTES DE USAR PELA PRIMEIRA VEZ:
      Rode .\Verificar-Atributos-AD.ps1 para confirmar que a matricula esta
      gravada no atributo EmployeeNumber do AD (ja e o padrao deste script).

    Formato da planilha (ex.: "desligados.xlsx"):
      - Colunas identificadas por LETRA, nao por nome de cabecalho:
          D = matricula (ex.: "EMP-90")
          E = nome do funcionario
          J = data de fim das ferias (so usada para linhas FERIAS)
          L = TIPO SITUACAO (AFASTADO / FERIAS / RESCISAO)
      - Por padrao so processa linhas RESCISAO e FERIAS (-SituacoesParaBloquear).
        AFASTADO fica de fora - rede de seguranca para nao bloquear afastamento.
      - Assume que a linha 1 e cabecalho e os dados comecam na linha 2.

    Uso:
      # Simulacao - desligamentos + ferias juntos, nada e alterado
      .\Bloqueio-Funcionarios-AD.ps1 -PlanilhaPath "C:\planilha.xlsx" -WhatIf

      # Execucao real, removendo tambem de um grupo (so quem foi desligado de vez)
      .\Bloqueio-Funcionarios-AD.ps1 -PlanilhaPath "C:\planilha.xlsx" -GrupoParaRemover "VPN-Acesso"

    O script:
      - Filtra as linhas cuja coluna L esta em -SituacoesParaBloquear (padrao: RESCISAO, FERIAS)
      - Localiza cada usuario no AD pela matricula (atributo -MatchAttribute)
      - Mostra uma tabela de conferencia e pede confirmacao antes de mexer em qualquer conta
      - Desabilita a conta (Disable-ADAccount)
      - Define a Description:
          - RESCISAO -> "Desligado dd/MM/aaaa - Retirar mensagem automática
            dd/MM/aaaa - Converter caixa em Sharedmailbox" (a segunda data e
            a primeira + 15 dias)
          - FERIAS (situacao em -SituacoesTemporarias) -> "Ferias ate dd/MM/aaaa",
            lida da coluna -ColunaDataFim. O Reativar-Ferias-AD.ps1 usa essa
            marca para reabilitar a conta automaticamente quando a data chegar.
      - -GrupoParaRemover so se aplica a contas PERMANENTES (nao mexe em quem
        esta so de ferias)
      - Gera um log em .\Logs e, se houver, planilhas com matriculas nao encontradas/ambiguas
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "High")]
param(
    # ============================================================
    # AJUSTE AQUI o caminho fixo da planilha no servidor.
    # ============================================================
    [string]$PlanilhaPath = "C:\BloqueioAD\desligados.xlsx",

    [ValidateSet("EmployeeID", "EmployeeNumber", "SamAccountName")]
    [string]$MatchAttribute = "EmployeeNumber",

    [string]$ColunaMatricula = "D",
    [string]$ColunaNome = "E",
    [string]$ColunaSituacao = "L",

    # Roda desligamento definitivo e bloqueio temporario de ferias na mesma execucao
    [string[]]$SituacoesParaBloquear = @("RESCISAO", "FERIAS"),

    # Quais dessas situacoes sao TEMPORARIAS (ferias) - para essas, a Description
    # vira "Ferias ate dd/MM/aa" (lida da coluna -ColunaDataFim) em vez de
    # "Desligado (dd/MM/aa)", e -GrupoParaRemover e ignorado (ferias nao remove
    # de grupo). O Reativar-Ferias-AD.ps1 usa essa marca para saber quando
    # reabilitar a conta.
    [string[]]$SituacoesTemporarias = @("FERIAS"),

    # Coluna com a data de retorno de ferias (ex.: "J" = DT_FIM_FERIAS)
    [string]$ColunaDataFim = "J",

    [int]$PrimeiraLinhaDados = 2,
    [int]$IndiceAba = 0,

    # Nome do grupo do AD do qual o usuario sera removido ao ser bloqueado (opcional)
    [string]$GrupoParaRemover,

    [string]$LogDir = "$PSScriptRoot\Logs",

    # Pula a pergunta de confirmacao (use com cuidado, ex.: em automacao ja validada)
    [switch]$Force
)

$ErrorActionPreference = "Stop"

Import-Module ActiveDirectory

if (-not (Test-Path $PlanilhaPath)) {
    throw "Planilha nao encontrada em: $PlanilhaPath"
}
if ([System.IO.Path]::GetExtension($PlanilhaPath).ToLowerInvariant() -ne ".xlsx") {
    throw "Este script le apenas .xlsx. Recebido: $PlanilhaPath"
}

if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $LogDir "Bloqueio.log"

function Write-Log {
    param([string]$Mensagem, [string]$Cor = "Gray")
    $linha = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Mensagem
    Write-Host $linha -ForegroundColor $Cor
    Add-Content -Path $logFile -Value $linha
}

Add-Content -Path $logFile -Value "`n----- Execucao $timestamp -----"

# ============================================================================
# Leitor XLSX autocontido (sem dependencia de modulo ImportExcel nem Excel
# instalado no servidor) - le o arquivo como pacote OOXML (zip + XML).
# ============================================================================
function Read-XlsxRows {
    param(
        [string]$Path,
        [int]$SheetIndex = 0
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        function Get-EntryXml($zip, $name) {
            $entry = $zip.Entries | Where-Object { $_.FullName -eq $name }
            if (-not $entry) { return $null }
            $sr = New-Object System.IO.StreamReader($entry.Open())
            $text = $sr.ReadToEnd()
            $sr.Close()
            return [xml]$text
        }
        function Get-RichText($node) {
            if ($null -eq $node) { return "" }
            if ($node -is [string]) { return $node }
            if ($node.InnerText) { return $node.InnerText }
            return [string]$node
        }

        # Mapeia indice/nome da aba -> arquivo sheetN.xml via workbook.xml + rels
        $workbookXml = Get-EntryXml $zip "xl/workbook.xml"
        $relsXml = Get-EntryXml $zip "xl/_rels/workbook.xml.rels"
        $sheets = @($workbookXml.workbook.sheets.sheet)
        if ($SheetIndex -ge $sheets.Count) {
            throw "A planilha nao tem uma aba de indice $SheetIndex (ela tem $($sheets.Count) aba(s))."
        }
        $rId = $sheets[$SheetIndex].id
        $target = ($relsXml.Relationships.Relationship | Where-Object { $_.Id -eq $rId }).Target
        $sheetPath = "xl/" + ($target -replace "^/?xl/", "")

        # Shared strings (texto das celulas)
        $shared = @()
        $sstXml = Get-EntryXml $zip "xl/sharedStrings.xml"
        if ($sstXml) {
            foreach ($si in $sstXml.sst.si) {
                if ($si.t) { $shared += (Get-RichText $si.t) }
                elseif ($si.r) { $shared += (($si.r | ForEach-Object { Get-RichText $_.t }) -join "") }
                else { $shared += "" }
            }
        }

        $sheetXml = Get-EntryXml $zip $sheetPath
        $rows = @($sheetXml.worksheet.sheetData.row)

        $resultado = @{}
        foreach ($row in $rows) {
            $numLinha = [int]$row.r
            $celulas = @{}
            foreach ($cell in $row.c) {
                if (-not $cell.r) { continue }
                $coluna = ($cell.r -replace '[0-9]', '')
                $valor = ""
                if ($cell.t -eq "s") {
                    if ($cell.v) { $valor = $shared[[int]$cell.v] }
                }
                elseif ($cell.t -eq "inlineStr") {
                    $valor = Get-RichText $cell.is.t
                }
                else {
                    $valor = $cell.v
                }
                $celulas[$coluna] = "$valor".Trim()
            }
            $resultado[$numLinha] = $celulas
        }
        return $resultado
    }
    finally {
        $zip.Dispose()
    }
}

# ============================================================================
# Carrega os dados da planilha em uma lista unica: NomePlanilha, Matricula
# ============================================================================
Write-Log "Lendo planilha: $PlanilhaPath (aba indice $IndiceAba)" "Cyan"
Write-Log "Colunas: matricula=$ColunaMatricula, nome=$ColunaNome, situacao=$ColunaSituacao | Situacoes que geram bloqueio: $($SituacoesParaBloquear -join ', ')" "Cyan"

$dados = @()
$linhasXlsx = Read-XlsxRows -Path $PlanilhaPath -SheetIndex $IndiceAba
$numerosLinha = $linhasXlsx.Keys | Sort-Object
$ignoradasPorSituacao = 0

foreach ($numLinha in $numerosLinha) {
    if ($numLinha -lt $PrimeiraLinhaDados) { continue }
    $celulas = $linhasXlsx[$numLinha]

    $matricula = $celulas[$ColunaMatricula]
    $nome = $celulas[$ColunaNome]
    $situacao = $celulas[$ColunaSituacao]

    if (-not $matricula -and -not $nome) { continue }

    if ($SituacoesParaBloquear.Count -gt 0 -and ($situacao -notin $SituacoesParaBloquear)) {
        $ignoradasPorSituacao++
        continue
    }

    $dataFim = $null
    if ($ColunaDataFim -and ($situacao -in $SituacoesTemporarias) -and $celulas[$ColunaDataFim]) {
        $bruto = $celulas[$ColunaDataFim]
        $numero = 0.0
        if ([double]::TryParse($bruto, [ref]$numero)) {
            $dataFim = [datetime]::FromOADate($numero)
        }
        else {
            Write-Log "AVISO: data de retorno invalida ('$bruto') para matricula '$matricula' (linha $numLinha) - tratando como sem data" "Yellow"
        }
    }

    $dados += [pscustomobject]@{ NomePlanilha = $nome; Matricula = $matricula; DataFim = $dataFim }
}

Write-Log "Linhas com situacao fora de '$($SituacoesParaBloquear -join ', ')' ignoradas: $ignoradasPorSituacao" "Cyan"
Write-Log "Total de funcionarios a processar (apos filtros): $($dados.Count)" "Cyan"

# ============================================================================
# Confere cada matricula no AD
# ============================================================================
$encontrados = @()
$naoEncontrados = @()
$ambiguos = @()
$jaDesabilitados = @()

foreach ($linha in $dados) {
    $nomePlanilha = $linha.NomePlanilha
    $matricula = $linha.Matricula

    if (-not $matricula) {
        Write-Log "Linha ignorada (matricula vazia): Nome='$nomePlanilha'" "Yellow"
        continue
    }

    $matriculaEscapada = $matricula -replace "'", "''"
    $filtro = "$MatchAttribute -eq '$matriculaEscapada'"
    $usuarios = @(Get-ADUser -Filter $filtro -Properties $MatchAttribute, Enabled, Description)

    if ($usuarios.Count -eq 0) {
        $naoEncontrados += [pscustomobject]@{ Nome = $nomePlanilha; Matricula = $matricula }
        Write-Log "NAO ENCONTRADO: Nome='$nomePlanilha' Matricula='$matricula'" "Yellow"
    }
    elseif ($usuarios.Count -gt 1) {
        $ambiguos += [pscustomobject]@{ Nome = $nomePlanilha; Matricula = $matricula; QtdEncontrada = $usuarios.Count }
        Write-Log "AMBIGUO ($($usuarios.Count) contas com a mesma matricula): Nome='$nomePlanilha' Matricula='$matricula'" "Yellow"
    }
    elseif (-not $usuarios[0].Enabled) {
        $jaDesabilitados += [pscustomobject]@{ Nome = $nomePlanilha; Matricula = $matricula; SamAccountName = $usuarios[0].SamAccountName }
        Write-Log "JA ESTAVA DESABILITADO: Nome='$nomePlanilha' Matricula='$matricula' SamAccountName='$($usuarios[0].SamAccountName)'" "DarkYellow"
    }
    else {
        $encontrados += [pscustomobject]@{
            NomePlanilha   = $nomePlanilha
            Matricula      = $matricula
            SamAccountName = $usuarios[0].SamAccountName
            NomeAD         = $usuarios[0].Name
            DN             = $usuarios[0].DistinguishedName
            DataFim        = $linha.DataFim
        }
    }
}

Write-Host ""
Write-Host "===== RESUMO DA CONFERENCIA =====" -ForegroundColor Cyan
Write-Host "Contas a bloquear agora......: $($encontrados.Count)"
Write-Host "Ja estavam desabilitadas.....: $($jaDesabilitados.Count)"
Write-Host "Nao encontradas no AD........: $($naoEncontrados.Count)"
Write-Host "Ambiguas (matricula repetida): $($ambiguos.Count)"
Write-Host ""

if ($encontrados.Count -gt 0) {
    Write-Host "Contas que serao bloqueadas:" -ForegroundColor Cyan
    $encontrados | Format-Table NomePlanilha, Matricula, SamAccountName, NomeAD, DataFim -AutoSize
}

if ($naoEncontrados.Count -gt 0) {
    $naoEncontrados | Export-Csv -Path (Join-Path $LogDir "NaoEncontrados_$timestamp.csv") -NoTypeInformation -Delimiter ";" -Encoding UTF8
    Write-Log "Lista de nao encontrados exportada para NaoEncontrados_$timestamp.csv" "Yellow"
}

if ($ambiguos.Count -gt 0) {
    $ambiguos | Export-Csv -Path (Join-Path $LogDir "Ambiguos_$timestamp.csv") -NoTypeInformation -Delimiter ";" -Encoding UTF8
    Write-Log "Lista de matriculas ambiguas exportada para Ambiguos_$timestamp.csv" "Yellow"
}

if ($encontrados.Count -eq 0) {
    Write-Log "Nenhuma conta valida para bloquear. Encerrando." "Yellow"
    return
}

if ($WhatIfPreference) {
    Write-Log "Modo -WhatIf: nenhuma conta foi alterada. Revise a tabela acima e rode sem -WhatIf para aplicar." "Cyan"
    return
}

if (-not $Force) {
    $resposta = Read-Host "Digite CONFIRMAR para bloquear as $($encontrados.Count) conta(s) acima no AD"
    if ($resposta -ne "CONFIRMAR") {
        Write-Log "Operacao cancelada pelo usuario." "Yellow"
        return
    }
}

$sucesso = 0
$falha = 0

foreach ($item in $encontrados) {
    if (-not $PSCmdlet.ShouldProcess($item.SamAccountName, "Disable-ADAccount")) { continue }

    try {
        Disable-ADAccount -Identity $item.DN

        if ($item.DataFim) {
            $descricao = "Ferias ate $($item.DataFim.ToString('dd/MM/yyyy'))"
        }
        else {
            $dataDesligamento = Get-Date
            $dataRetirarMensagem = $dataDesligamento.AddDays(15)
            $descricao = "Desligado $($dataDesligamento.ToString('dd/MM/yyyy')) - Retirar mensagem automática $($dataRetirarMensagem.ToString('dd/MM/yyyy')) - Converter caixa em Sharedmailbox"
        }
        Set-ADUser -Identity $item.DN -Description $descricao

        if ($GrupoParaRemover -and -not $item.DataFim) {
            try {
                Remove-ADGroupMember -Identity $GrupoParaRemover -Members $item.SamAccountName -Confirm:$false
                Write-Log "  Removido do grupo '$GrupoParaRemover'" "Green"
            }
            catch {
                Write-Log "  AVISO: nao foi possivel remover do grupo '$GrupoParaRemover': $($_.Exception.Message)" "Yellow"
            }
        }

        Write-Log "BLOQUEADO: $($item.SamAccountName) ($($item.NomeAD))" "Green"
        $sucesso++
    }
    catch {
        Write-Log "ERRO ao bloquear $($item.SamAccountName): $($_.Exception.Message)" "Red"
        $falha++
    }
}

Write-Host ""
Write-Log "Concluido. Bloqueados: $sucesso | Falhas: $falha | Log completo em: $logFile" "Cyan"
