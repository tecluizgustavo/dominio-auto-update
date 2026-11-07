# dominio-auto-update

Automação em PowerShell que aplica sozinha as atualizações do **Domínio Contábil** (Thomson Reuters) no servidor, sem ninguém conectado, e envia um relatório por e-mail via Microsoft Graph.

Ela automatiza o procedimento manual que um analista fazia: parar o agente e o serviço, rodar o instalador, atualizar o banco pelo cliente e subir tudo de volta.

> O instalador do Domínio **não aceita instalação silenciosa** (a central de suporte confirma). Por isso a automação clica nas telas do instalador por mensagens do Windows.

## O que faz, em ordem

1. Procura o pacote `C*.exe` mais novo na pasta onde o Gerenciador de Atualizações do Domínio baixa as atualizações.
2. Compara com o último pacote aplicado (arquivo de estado). Sem pacote novo, sai sem fazer nada.
3. Aborta se o pacote foi alterado há menos de 10 minutos (pode estar baixando) ou se há estações conectadas ao banco.
4. Para o `Agente_comunicacao` e o serviço `ServicoDominioAtendimento`.
5. Executa o instalador (Inno Setup) e clica em Avançar, Atualizar e Concluir.
6. Sobe o serviço, abre o `contabil.exe`, faz login com a credencial salva e conduz a atualização do banco (Sim, caixa de ciência, Iniciar Atualização, OK).
7. Fecha o cliente, sobe o agente e fecha o Domínio Folha, se estiver aberto.
8. Grava o log, atualiza o estado e envia o relatório com o log em anexo.

Se a fase do banco falhar, o script fecha o cliente, sobe o agente de volta e **não** grava o estado como aplicado.

## Arquivos

| Arquivo | Para quê |
|---|---|
| `scripts/Atualiza-Dominio.ps1` | Script principal. Todas as configurações ficam no bloco `$Cfg` no topo. |
| `scripts/Salvar-Credencial.ps1` | Guarda usuário e senha criptografados (DPAPI do Windows). |
| `scripts/Agendar-Tarefa.ps1` | Cria a tarefa agendada, que roda com ou sem usuário conectado. |
| `scripts/Diag-Instalador.ps1` | Diagnóstico: lista o que o Windows enxerga numa janela de instalador. |

## Requisitos

- Windows Server com PowerShell 5.1, rodando elevado.
- Usuário administrador local com o direito "Fazer logon como tarefa em lote".
- Domínio Contábil com o Gerenciador de Atualizações baixando os pacotes em disco.
- Para o relatório por e-mail: um app registration no Microsoft Entra ID com permissão de **aplicativo** `Mail.Send`.

## Instalação

1. Copie a pasta `scripts` para o servidor, por exemplo `D:\DOMINIO\AtualizaDomSis\`.
2. Ajuste o bloco `$Cfg` do `Atualiza-Dominio.ps1`:
   - `PackageRoot`, `ClientExe`, `ClientProcess`, `AgentProcess`, `ServiceMatch`.
   - `DbPort`: porta do banco, para detectar estações conectadas.
   - `FolhaMatch`: nome do processo do Domínio Folha.
3. Marque a versão instalada como aplicada (uma vez):

   ```powershell
   .\Atualiza-Dominio.ps1 -Seed
   ```
4. Salve a credencial do login do Domínio:

   ```powershell
   .\Salvar-Credencial.ps1
   ```
5. Confira o que ele faria, sem alterar nada:

   ```powershell
   .\Atualiza-Dominio.ps1 -DryRun
   ```
6. Crie a tarefa agendada (pede a senha do Windows do usuário, que o Agendador guarda):

   ```powershell
   .\Agendar-Tarefa.ps1 -Hora 05:00
   ```

O horário deve ficar depois da janela de download do Gerenciador de Atualizações e antes do expediente.

## E-mail via Microsoft Graph

Preencha no `$Cfg`: `GraphTenantId`, `GraphClientId`, `MailFrom` (caixa que envia) e `MailTo`. O client secret **não** vai no script:

```powershell
.\Salvar-Credencial.ps1 -Path D:\DOMINIO\AtualizaDomSis\cred-graph.xml -Mensagem "Client ID e secret"
```

No prompt, o usuário é o Client ID e a senha é o client secret. Teste sem atualizar nada:

```powershell
.\Atualiza-Dominio.ps1 -TestEmail
```

O relatório só sai quando o script agiu ou deu erro, então os dias sem pacote novo não geram e-mail. Há também um modo SMTP alternativo, usado se `GraphTenantId` estiver vazio.

## Como funciona por dentro

- **Cliques por Win32.** O instalador é um Inno Setup, e o cliente é PowerBuilder. Os controles não expõem `InvokePattern` no UI Automation, então o script usa `EnumWindows`, `BM_CLICK` e `WM_SETTEXT`.
- **Sem usuário conectado.** Na tarefa agendada ("executar estando o usuário conectado ou não") tudo roda na sessão 0, e as janelas têm `WS_VISIBLE` desligado. O script aceita janelas invisíveis e mede a visibilidade dos controles relativa à janela de topo, para distinguir a página atual do assistente.
- **Filtros de botão ancorados.** `Avançar` casaria também com os botões da barra de rolagem da tela de informações, então o filtro exige o nome exato.
- **Credenciais.** O login do Domínio fica em um arquivo `Export-Clixml`, criptografado com DPAPI: só o mesmo usuário do Windows, na mesma máquina, consegue lê-lo. O script tenta o login no máximo 2 vezes, para não bloquear a conta.
- **Agente fora da árvore da tarefa.** O agente é iniciado por `Win32_Process.Create`; senão o Agendador o encerra quando o script termina.

## Limitações e cuidados

- A tela "deseja atualizar o banco?" só aparece quando há atualização real do banco. Ela foi validada com o usuário conectado, mas na sessão 0 só será confirmada na primeira atualização real. Se algo falhar ali, o log traz as janelas e botões vistos.
- Se a senha do Windows do usuário da tarefa mudar, recrie a tarefa com `Agendar-Tarefa.ps1`. Se a senha do login do Domínio mudar, rode `Salvar-Credencial.ps1` de novo.
- A automação marca sozinha a caixa "ciência de que posso ficar sem acesso se não houver backup recente". Mantenha o backup do Domínio funcionando.
- Interface do instalador ou do cliente mudando em versões futuras pode exigir ajuste nos filtros de botão.

## Segurança

Este repositório não contém credenciais, IDs de tenant ou endereços internos. Os arquivos `cred-*.xml`, os logs e o estado estão no `.gitignore`. Nunca coloque o client secret direto no script.
