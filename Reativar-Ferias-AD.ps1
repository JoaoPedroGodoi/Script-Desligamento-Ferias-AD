<#
    Reativar-Ferias-AD.ps1

    Reabilita no Active Directory as contas que foram bloqueadas
    temporariamente por ferias (via Bloqueio-Funcionarios-AD.ps1 com
    -ColunaDataFim) e cuja data de retorno ja chegou.

    Como funciona:
      - Nao le a planilha. Procura direto no AD por contas DESABILITADAS cuja
        Description bate com o padrao "Ferias ate dd/MM/aaaa" (gravado pelo
        Bloqueio-Funcionarios-AD.ps1; tambem aceita o formato antigo de 2
        digitos no ano, dd/MM/aa).
      - Se a data de retorno for hoje ou ja tiver passado, reabilita a conta
        (Enable-ADAccount) e limpa a Description.
      - Contas com data de retorno futura sao listadas mas nao mexidas.
      - Se nenhuma conta for encontrada, mostra um diagnostico (quantas contas
        desabilitadas existem no total e quantas tem "ferias" na Description
        sem bater no padrao exato) para ajudar a achar o motivo.

    Uso:
      # Simulacao - mostra quem seria reabilitado, sem alterar nada
      .\Reativar-Ferias-AD.ps1 -WhatIf

      # Execucao real (pede confirmacao antes de aplicar)
      .\Reativar-Ferias-AD.ps1

    Pensado para rodar periodicamente (ex.: 1x por dia via Agendador de
    Tarefas do Windows) com -Force, ja que so mexe em contas cuja data de
    retorno realmente chegou.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "High")]
param(
    [string]$LogDir = "$PSScriptRoot\Logs",

    # Pula a pergunta de confirmacao (use com cuidado, ex.: em automacao ja validada)
    [switch]$Force
)

$ErrorActionPreference = "Stop"

Import-Module ActiveDirectory

if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $LogDir "Reativacao.log"

function Write-Log {
    param([string]$Mensagem, [string]$Cor = "Gray")
    $linha = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Mensagem
    Write-Host $linha -ForegroundColor $Cor
    Add-Content -Path $logFile -Value $linha
}

Add-Content -Path $logFile -Value "`n----- Execucao $timestamp -----"

function ConvertFrom-DataDescricao {
    param([string]$Texto)
    foreach ($formato in @("dd/MM/yyyy", "dd/MM/yy")) {
        try { return [datetime]::ParseExact($Texto, $formato, $null) } catch { }
    }
    return $null
}

$padraoFerias = '^Ferias ate (\d{2}/\d{2}/\d{2,4})$'
$todosDesabilitados = @(Get-ADUser -Filter { Enabled -eq $false } -Properties Description, DistinguishedName)
$candidatos = $todosDesabilitados | Where-Object { $_.Description -match $padraoFerias }

Write-Log "Contas desabilitadas no AD: $($todosDesabilitados.Count) | Com marca de ferias reconhecida: $($candidatos.Count)" "Cyan"

if ($candidatos.Count -eq 0) {
    $comFeriasNoTexto = @($todosDesabilitados | Where-Object { $_.Description -match "(?i)ferias" -and $_.Description -notmatch $padraoFerias })
    if ($comFeriasNoTexto.Count -gt 0) {
        Write-Log "AVISO: $($comFeriasNoTexto.Count) conta(s) desabilitada(s) tem 'ferias' na Description mas NAO bateram no padrao esperado 'Ferias ate dd/MM/aaaa'. Description encontrada (para conferir formatacao):" "Yellow"
        $comFeriasNoTexto | ForEach-Object { Write-Log "  $($_.SamAccountName): '$($_.Description)'" "Yellow" }
    }
}

$hoje = (Get-Date).Date
$paraReativar = @()
$aindaEmFerias = @()

foreach ($usuario in $candidatos) {
    $null = $usuario.Description -match $padraoFerias
    $dataRetorno = ConvertFrom-DataDescricao $Matches[1]
    if (-not $dataRetorno) {
        Write-Log "AVISO: nao consegui converter a data '$($Matches[1])' da conta $($usuario.SamAccountName) - pulando" "Yellow"
        continue
    }

    if ($dataRetorno -le $hoje) {
        $paraReativar += [pscustomobject]@{
            SamAccountName = $usuario.SamAccountName
            Nome           = $usuario.Name
            DataRetorno    = $dataRetorno
            DN             = $usuario.DistinguishedName
        }
    }
    else {
        $aindaEmFerias += [pscustomobject]@{
            SamAccountName = $usuario.SamAccountName
            Nome           = $usuario.Name
            DataRetorno    = $dataRetorno
        }
    }
}

Write-Host ""
Write-Host "===== RESUMO =====" -ForegroundColor Cyan
Write-Host "Contas a reativar agora (retorno <= hoje): $($paraReativar.Count)"
Write-Host "Ainda em ferias (retorno futuro)..........: $($aindaEmFerias.Count)"
Write-Host ""

if ($paraReativar.Count -gt 0) {
    Write-Host "Contas que serao reativadas:" -ForegroundColor Cyan
    $paraReativar | Format-Table SamAccountName, Nome, DataRetorno -AutoSize
}
if ($aindaEmFerias.Count -gt 0) {
    Write-Host "Ainda em ferias (nada sera feito):" -ForegroundColor DarkGray
    $aindaEmFerias | Format-Table SamAccountName, Nome, DataRetorno -AutoSize
}

if ($paraReativar.Count -eq 0) {
    Write-Log "Nenhuma conta para reativar hoje." "Yellow"
    return
}

if ($WhatIfPreference) {
    Write-Log "Modo -WhatIf: nenhuma conta foi alterada." "Cyan"
    return
}

if (-not $Force) {
    $resposta = Read-Host "Digite CONFIRMAR para reativar as $($paraReativar.Count) conta(s) acima no AD"
    if ($resposta -ne "CONFIRMAR") {
        Write-Log "Operacao cancelada pelo usuario." "Yellow"
        return
    }
}

$sucesso = 0
$falha = 0

foreach ($item in $paraReativar) {
    if (-not $PSCmdlet.ShouldProcess($item.SamAccountName, "Enable-ADAccount")) { continue }

    try {
        Enable-ADAccount -Identity $item.DN
        Set-ADUser -Identity $item.DN -Description $null
        Write-Log "REATIVADO: $($item.SamAccountName) ($($item.Nome))" "Green"
        $sucesso++
    }
    catch {
        Write-Log "ERRO ao reativar $($item.SamAccountName): $($_.Exception.Message)" "Red"
        $falha++
    }
}

Write-Host ""
Write-Log "Concluido. Reativados: $sucesso | Falhas: $falha | Log completo em: $logFile" "Cyan"
