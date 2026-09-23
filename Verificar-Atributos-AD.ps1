<#
    Verificar-Atributos-AD.ps1

    Uso: rode este script ANTES do Bloqueio-Funcionarios-AD.ps1, uma unica vez,
    para descobrir em qual atributo do Active Directory a matricula do
    funcionario esta gravada (EmployeeID, EmployeeNumber, ou nenhum dos dois).

    Execute no servidor onde o modulo ActiveDirectory do PowerShell esta
    disponivel (normalmente o proprio DC, ou uma maquina com RSAT-AD-PowerShell).

    Parametros:
      -Quantidade   Numero de usuarios de exemplo a listar (padrao: 15)
      -Matricula    Se informado, procura diretamente por essa matricula em
                    varios atributos e mostra em qual(is) ela foi encontrada.

    Exemplos:
      .\Verificar-Atributos-AD.ps1
      .\Verificar-Atributos-AD.ps1 -Quantidade 30
      .\Verificar-Atributos-AD.ps1 -Matricula "12345"
#>

[CmdletBinding()]
param(
    [int]$Quantidade = 15,
    [string]$Matricula
)

Import-Module ActiveDirectory -ErrorAction Stop

if ($Matricula) {
    Write-Host "Procurando a matricula '$Matricula' nos atributos EmployeeID, EmployeeNumber e SamAccountName..." -ForegroundColor Cyan
    Write-Host ""

    $porEmployeeID = Get-ADUser -Filter "EmployeeID -eq '$Matricula'" -Properties EmployeeID, EmployeeNumber, SamAccountName, Enabled
    $porEmployeeNumber = Get-ADUser -Filter "EmployeeNumber -eq '$Matricula'" -Properties EmployeeID, EmployeeNumber, SamAccountName, Enabled
    $porSam = Get-ADUser -Filter "SamAccountName -eq '$Matricula'" -Properties EmployeeID, EmployeeNumber, SamAccountName, Enabled

    if ($porEmployeeID) { Write-Host "Encontrado via EmployeeID:" -ForegroundColor Green; $porEmployeeID | Format-Table Name, SamAccountName, EmployeeID, EmployeeNumber, Enabled -AutoSize }
    if ($porEmployeeNumber) { Write-Host "Encontrado via EmployeeNumber:" -ForegroundColor Green; $porEmployeeNumber | Format-Table Name, SamAccountName, EmployeeID, EmployeeNumber, Enabled -AutoSize }
    if ($porSam) { Write-Host "Encontrado via SamAccountName:" -ForegroundColor Green; $porSam | Format-Table Name, SamAccountName, EmployeeID, EmployeeNumber, Enabled -AutoSize }

    if (-not $porEmployeeID -and -not $porEmployeeNumber -and -not $porSam) {
        Write-Host "Nenhum usuario encontrado com essa matricula em nenhum dos tres atributos." -ForegroundColor Yellow
    }

    return
}

Write-Host "Listando $Quantidade usuarios habilitados como amostra (Name, SamAccountName, EmployeeID, EmployeeNumber, Description):" -ForegroundColor Cyan
Write-Host ""

Get-ADUser -Filter "Enabled -eq 'True'" -Properties EmployeeID, EmployeeNumber, Description |
    Select-Object -First $Quantidade Name, SamAccountName, EmployeeID, EmployeeNumber, Description |
    Format-Table -AutoSize

Write-Host ""
Write-Host "Olhe a saida acima:" -ForegroundColor Cyan
Write-Host " - Se a coluna EmployeeID tiver a matricula -> use -MatchAttribute EmployeeID no script principal (padrao)."
Write-Host " - Se a coluna EmployeeNumber tiver a matricula -> use -MatchAttribute EmployeeNumber."
Write-Host " - Se nenhuma tiver, mas o SamAccountName (login) for a propria matricula -> use -MatchAttribute SamAccountName."
Write-Host " - Se nao for nenhum desses, rode de novo com -Matricula '<uma matricula real>' para conferir onde ela esta."
