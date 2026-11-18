unit uPrincipal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.StdCtrls, System.Win.ComObj, Winapi.TlHelp32;

type
  TfrmPrincipal = class(TForm)
    lblStatus: TLabel;
    lblPID: TLabel;
    edtPID: TEdit;
    btnFecharPorPID: TButton;
    btnExcelPID: TButton;
    btnFecharNotepad: TButton;
    procedure btnFecharPorPIDClick(Sender: TObject);
    procedure btnExcelPIDClick(Sender: TObject);
    procedure btnFecharNotepadClick(Sender: TObject);
  private
    // Método 1: Função base por PID direto
    function FinalizarProcessoPorPID(APID: Cardinal): Boolean;
    // Método 2: Automação OLE (Excel/Word) usando HWND -> PID
    procedure ExportarExcelComSeguranca(ASimularFalha: Boolean = True);
    // Método 3: Varredura de processos pelo nome (.exe)
    function FinalizarProcessoPorNome(const ANomeExecutavel: string): Integer;
  public
  end;

var
  frmPrincipal: TfrmPrincipal;

implementation

{$R *.dfm}

procedure TfrmPrincipal.btnFecharPorPIDClick(Sender: TObject);
var
  lPID: Cardinal;
begin
  lPID := StrToIntDef(Trim(edtPID.Text), 0);
  if lPID <= 4 then
  begin
    lblStatus.Caption := 'Informe um PID válido (maior que 4) obtido no tasklist.';
    edtPID.SetFocus;
    Exit;
  end;

  if FinalizarProcessoPorPID(lPID) then
    lblStatus.Caption := Format('Método 1: Processo PID %d finalizado com sucesso.', [lPID])
  else
    lblStatus.Caption := Format('Falha ao finalizar PID %d (inexistente, protegido ou sem permissão).', [lPID]);
end;

procedure TfrmPrincipal.btnExcelPIDClick(Sender: TObject);
begin
  ExportarExcelComSeguranca(True);
end;

procedure TfrmPrincipal.btnFecharNotepadClick(Sender: TObject);
var
  lQtdEncerrada: Integer;
begin
  lQtdEncerrada := FinalizarProcessoPorNome('notepad.exe');
  if lQtdEncerrada > 0 then
    lblStatus.Caption := Format('Método 3: %d processo(s) notepad.exe finalizado(s).', [lQtdEncerrada])
  else
    lblStatus.Caption := 'Nenhum processo notepad.exe estava aberto no momento.';
end;

// =========================================================================
// MÉTODO 1: A Função Base para Finalizar Processo por PID
// =========================================================================
function TfrmPrincipal.FinalizarProcessoPorPID(APID: Cardinal): Boolean;
var
  lProcessHandle: THandle;
begin
  Result := False;

  // Proteção: bloqueia processos críticos do sistema (0 e 4) e a própria aplicação
  if (APID <= 4) or (APID = GetCurrentProcessId) then
    Exit;

  lProcessHandle := OpenProcess(PROCESS_TERMINATE, False, APID);
  if lProcessHandle <> 0 then
  begin
    try
      Result := TerminateProcess(lProcessHandle, 0);
    finally
      // Obrigatório fechar o Handle para não vazar recursos do Windows
      CloseHandle(lProcessHandle);
    end;
  end;
end;

// =========================================================================
// MÉTODO 2: Finalizar Processo Windows no Delphi por HWND (Excel e Word)
// =========================================================================
procedure TfrmPrincipal.ExportarExcelComSeguranca(ASimularFalha: Boolean);
var
  loPlanilha: OleVariant;
  lHWND: HWND;
  lProcID: Cardinal;
  lSucesso: Boolean;
begin
  lProcID := 0;
  lSucesso := False;
  lblStatus.Caption := 'Iniciando Excel via OLE (mantenha outra planilha aberta para testar)...';
  Application.ProcessMessages;

  // Inicia uma instância isolada do Excel via OLE
  loPlanilha := CreateOleObject('Excel.Application');
  try
    // 1. Captura o Handle da janela e extrai o PID exclusivo desta instância
    lHWND := loPlanilha.HWND;
    GetWindowThreadProcessId(lHWND, lProcID);

    lblStatus.Caption := Format('Excel criado (PID: %d). Simulando falha em 2 segundos...', [lProcID]);
    Application.ProcessMessages;

    loPlanilha.Visible := False;
    loPlanilha.DisplayAlerts := False;
    loPlanilha.Workbooks.Add;
    Sleep(2000);

    // Simula um erro no meio da rotina antes de chegar no Quit
    if ASimularFalha then
      raise Exception.Create('Erro simulado no meio da exportação!');

    // 2. Tenta o encerramento normal primeiro
    loPlanilha.Quit;
    lSucesso := True;
    lblStatus.Caption := Format('Método 2: Finalizado de forma limpa (PID %d encerrado).', [lProcID]);
  finally
    // 3. Desvincula a referência COM ANTES de forçar o encerramento
    loPlanilha := Unassigned;

    // 4. Rede de segurança: se deu erro antes do Quit, mata APENAS este PID
    if (not lSucesso) and (lProcID > 0) then
    begin
      if FinalizarProcessoPorPID(lProcID) then
        lblStatus.Caption := Format('Método 2: Falha capturada! Apenas o PID %d foi encerrado.', [lProcID]);
    end;
  end;
end;

// =========================================================================
// MÉTODO 3: Como Matar Processo pelo Nome do Executável (.exe)
// =========================================================================
function TfrmPrincipal.FinalizarProcessoPorNome(const ANomeExecutavel: string): Integer;
var
  lSnapshot: THandle;
  lProcessEntry: TProcessEntry32;
  lContinuar: Boolean;
begin
  Result := 0;

  lSnapshot := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  if lSnapshot = INVALID_HANDLE_VALUE then
    Exit;

  try
    lProcessEntry.dwSize := SizeOf(TProcessEntry32);
    lContinuar := Process32First(lSnapshot, lProcessEntry);

    while lContinuar do
    begin
      if SameText(ExtractFileName(lProcessEntry.szExeFile), ANomeExecutavel) then
      begin
        // Reutiliza a função segura do Método 1 (já protegida contra auto-encerramento)
        if FinalizarProcessoPorPID(lProcessEntry.th32ProcessID) then
          Inc(Result);
      end;
      lContinuar := Process32Next(lSnapshot, lProcessEntry);
    end;
  finally
    CloseHandle(lSnapshot);
  end;
end;

end.

