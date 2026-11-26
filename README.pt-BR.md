🌐 **Idioma / Language:** [Português (Brasil)](https://github.com) | [English](https://github.com)

# Finalizar Processo Windows no Delphi (3 Métodos Seguros via WinAPI)

Projeto VCL prático e compilável demonstrando como encerrar processos no Windows pelo **Delphi** de forma segura, sem vazar recursos (*Handle Leaks*) e sem derrubar arquivos que o usuário já esteja editando na máquina.

> 📖 **Tutorial completo com explicação linha por linha e dicas de prevenção:**  
> [Finalizar Processo Windows no Delphi: 3 Métodos (PID e .EXE)](https://programacaolivre.com.br/finalizar-processo-windows-no-delphi/)

---

## 🎯 O Que Este Projeto Resolve

1. **Encerramento Cirúrgico por PID (`FinalizarProcessoPorPID`):**
   * Solicita permissão (`PROCESS_TERMINATE`), finaliza um processo específico pelo identificador numérico (`PID`) e libera o *Handle* imediatamente com `CloseHandle` no bloco `finally`.
   * Inclui trava de segurança contra encerramento acidental de processos críticos do sistema (`PID <= 4`) ou da própria aplicação (`GetCurrentProcessId`).

2. **Automação OLE Segura para Excel e Word (`ExportarExcelComSeguranca`):**
   * Captura o `HWND` exclusivo da janela criada via `CreateOleObject` e extrai seu `PID` com `GetWindowThreadProcessId`.
   * Limpa a referência COM (`Unassigned`) **antes** de encerrar o processo no `finally` (evitando exceções de *Servidor RPC indisponível*) e derruba **apenas** a instância oculta criada pelo sistema, preservando planilhas abertas pelo usuário.

3. **Encerramento em Lote por Nome (`FinalizarProcessoPorNome`):**
   * Utiliza `CreateToolhelp32Snapshot` (`Winapi.TlHelp32`) para percorrer os processos ativos do Windows e encerrar todas as instâncias de um executável específico (`.exe`) reaproveitando a função segura do Método 1.

---

## 🛠️ Compatibilidade & Padrão de Código

* **Desenvolvido e testado em:** Embarcadero Delphi Community Edition (RAD Studio 11+).
* **Retrocompatível com:** Delphi 10 Seattle, Delphi XE e versões anteriores (código escrito em sintaxe clássica, sem *inline variables* `var x :=` e zero *warnings* de compilação).
* **Plataforma:** Windows (Win32 / Win64).

---

## 🚀 Como Testar na Prática

1. Clone ou baixe este repositório:
   ```bash
   git clone https://github.com
   ```
2. Abra o arquivo `.dproj` no seu Delphi (`F11` / `Ctrl + F9`) e execute (`F9`).
3. **Para testar o Método 1 (PID):**
   * Abra duas janelas diferentes do Bloco de Notas (`notepad.exe`).
   * No Prompt de Comando (CMD), execute:
     ```bat
     tasklist /FI "IMAGENAME eq notepad.exe"
     ```
   * Digite apenas um dos PIDs retornados no campo da aplicação e clique em **Finalizar por PID Direto**. Apenas a janela correspondente será fechada.
4. **Para testar o Método 2 (Excel/OLE):**
   * Deixe uma planilha do Microsoft Excel aberta na sua área de trabalho.
   * Clique em **Testar Excel/OLE** — o projeto criará uma instância oculta, simulará uma exceção no meio da rotina e encerrará apenas o processo oculto, mantendo sua planilha aberta intacta.

---

## 📂 Estrutura dos Arquivos

* `uPrincipal.pas` / `uPrincipal.dfm` — Formulário VCL e implementação completa dos 3 métodos comentados.

---

## 📄 Licença

Código disponibilizado sob licença MIT para fins educacionais e uso livre em projetos pessoais ou comerciais pelo hub **[Programação Livre](https://programacaolivre.com.br)**.
