# Script Desligamento e Férias AD

Automação em PowerShell para bloqueio de acesso no Active Directory a partir de uma planilha de RH: desligamentos definitivos e bloqueio temporário de quem está de férias, com reativação automática no retorno.

Não depende de módulos externos (ImportExcel, Excel instalado) — lê o `.xlsx` diretamente como pacote OOXML. Só precisa do módulo `ActiveDirectory` do PowerShell, já disponível em qualquer servidor com a role AD DS ou RSAT-AD-PowerShell.

## Impacto

Em produção no servidor Active Directory da TEJOFRAN, operado pelo time de TI a cada ciclo de desligamento/férias reportado pelo RH.

- **Tempo**: o bloqueio manual (buscar a conta no AD, desabilitar, ajustar a descrição) levava cerca de **15 minutos por funcionário**. Com o volume atual de **20 a 30 funcionários/mês**, isso representava de **5 a 7,5 horas de trabalho manual repetitivo por mês** — hoje é uma única execução em lote, na casa de minutos, independente do volume.
- **Volume**: processa em lote os **20 a 30 desligamentos/férias por mês** direto da planilha que o RH já envia, sem digitação manual conta por conta.
- **Risco**: elimina o bloqueio manual conta-por-conta — reduz a chance de esquecer um desligamento, bloquear a pessoa errada (nomes parecidos) ou deixar de reativar alguém que voltou de férias na data certa. Matrículas não encontradas ou ambíguas nunca são bloqueadas automaticamente, ficam sinalizadas para conferência.

## Requisitos

- Windows Server com o módulo PowerShell `ActiveDirectory` (RSAT-AD-PowerShell)
- Executar como administrador, com permissão para desabilitar/reabilitar contas no AD
- PowerShell 5.1+

## Scripts

| Script | O que faz |
|---|---|
| [`Verificar-Atributos-AD.ps1`](Verificar-Atributos-AD.ps1) | Diagnóstico: descobre em qual atributo do AD (`EmployeeID`, `EmployeeNumber` ou `SamAccountName`) a matrícula do funcionário está gravada. Rodar uma vez, antes de tudo. |
| [`Bloqueio-Funcionarios-AD.ps1`](Bloqueio-Funcionarios-AD.ps1) | Lê a planilha e, numa única execução, desabilita quem foi desligado de vez e quem saiu de férias (bloqueio temporário). |
| [`Reativar-Ferias-AD.ps1`](Reativar-Ferias-AD.ps1) | Varre o AD por contas desabilitadas temporariamente por férias cuja data de retorno já chegou e reabilita automaticamente. Pensado para rodar 1x/dia via Agendador de Tarefas. |

## Configuração inicial

1. **Descubra o atributo certo da matrícula**, rodando no servidor:
   ```powershell
   .\Verificar-Atributos-AD.ps1
   ```
   Isso lista uma amostra de usuários com `EmployeeID` e `EmployeeNumber` lado a lado, para você comparar com a matrícula da sua planilha. Se preferir testar uma matrícula específica:
   ```powershell
   .\Verificar-Atributos-AD.ps1 -Matricula "EMP-12345"
   ```

2. **Ajuste o caminho fixo da planilha** em [`Bloqueio-Funcionarios-AD.ps1`](Bloqueio-Funcionarios-AD.ps1), logo no início do bloco `param(...)`:
   ```powershell
   [string]$PlanilhaPath = "C:\BloqueioAD\desligados.xlsx",
   ```

3. **Confira as colunas da sua planilha** — por padrão o script espera (identificando por letra, não por nome de cabeçalho):

   | Coluna | Conteúdo |
   |---|---|
   | D | Matrícula do funcionário |
   | E | Nome do funcionário |
   | J | Data de fim das férias (só relevante para linhas de férias) |
   | L | Tipo de situação (`RESCISAO`, `FERIAS`, `AFASTADO`, ...) |

   Se a sua planilha usar outras colunas, todos os parâmetros são ajustáveis (`-ColunaMatricula`, `-ColunaNome`, `-ColunaSituacao`, `-ColunaDataFim`).

## Uso

Sempre teste com `-WhatIf` primeiro — ele mostra exatamente o que seria feito, sem alterar nada no AD:

```powershell
.\Bloqueio-Funcionarios-AD.ps1 -WhatIf
```

Execução real (pede confirmação digitada antes de aplicar qualquer bloqueio):

```powershell
.\Bloqueio-Funcionarios-AD.ps1
```

Desligamento definitivo, removendo também de um grupo do AD (ex.: VPN):

```powershell
.\Bloqueio-Funcionarios-AD.ps1 -GrupoParaRemover "VPN-Acesso"
```

Reativação de quem voltou de férias:

```powershell
.\Reativar-Ferias-AD.ps1 -WhatIf
.\Reativar-Ferias-AD.ps1
```

Para automatizar a reativação diária via Agendador de Tarefas do Windows, use `-Force` (pula a confirmação interativa) na ação da tarefa:

```
powershell.exe -ExecutionPolicy Bypass -File "C:\caminho\Reativar-Ferias-AD.ps1" -Force
```

## Como o bloqueio funciona

Por padrão, uma única execução do `Bloqueio-Funcionarios-AD.ps1` processa dois tipos de situação (`-SituacoesParaBloquear`, padrão `RESCISAO, FERIAS`):

- **Desligamento definitivo (`RESCISAO`)**: desabilita a conta e grava na `Description`:
  `Desligado dd/MM/yyyy - Retirar mensagem automática dd/MM/yyyy - Converter caixa em Sharedmailbox` (a segunda data é a primeira + 15 dias). Se `-GrupoParaRemover` for informado, também remove o usuário desse grupo.

- **Férias (`FERIAS`, temporário)**: desabilita a conta e grava `Ferias ate dd/MM/yyyy`, usando a data da coluna `-ColunaDataFim`. Essa marca na `Description` é o que o `Reativar-Ferias-AD.ps1` procura depois para saber quando reabilitar — nunca remove de grupos.

`AFASTADO` fica de fora por padrão — rede de segurança para não bloquear quem só está afastado (atestado, licença, etc.), já que a planilha de RH costuma trazer os três tipos misturados na mesma aba.

## Segurança

- Toda execução mostra uma tabela de conferência (matrícula, nome, conta do AD encontrada) e exige confirmação digitada antes de alterar qualquer conta — a menos que `-Force` seja usado
- `-WhatIf` simula sem tocar em nada
- Matrículas não encontradas no AD ou ambíguas (mais de uma conta com a mesma matrícula) nunca são bloqueadas automaticamente — ficam listadas à parte em `Logs\`
- Logs de todas as execuções ficam em `Logs\Bloqueio.log` e `Logs\Reativacao.log`

**A pasta `Logs\` e qualquer planilha `.xlsx`/`.csv` são ignoradas pelo git** (veja [`.gitignore`](.gitignore)) — elas contêm dados reais de funcionários e nunca devem ser commitadas.
