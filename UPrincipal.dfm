object frmPrincipal: TfrmPrincipal
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = 'Programa'#231#227'o Livre - Finalizar Processos no Delphi (3 M'#233'todos)'
  ClientHeight = 245
  ClientWidth = 470
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  Position = poScreenCenter
  TextHeight = 15
  object lblStatus: TLabel
    Left = 24
    Top = 16
    Width = 422
    Height = 41
    AutoSize = False
    Caption = 
      'Escolha um dos 3 m'#233'todos abaixo para testar no Gerenciador de Ta' +
      'refas:'
    WordWrap = True
  end
  object lblPID: TLabel
    Left = 24
    Top = 68
    Width = 174
    Height = 15
    Caption = 'M'#233'todo 1 '#8221' Digite o PID (tasklist):'
  end
  object edtPID: TEdit
    Left = 24
    Top = 89
    Width = 200
    Height = 23
    NumbersOnly = True
    TabOrder = 0
    TextHint = 'Ex: 14820'
  end
  object btnFecharPorPID: TButton
    Left = 241
    Top = 80
    Width = 205
    Height = 38
    Caption = '1. Finalizar por PID Direto'
    TabOrder = 1
    OnClick = btnFecharPorPIDClick
  end
  object btnExcelPID: TButton
    Left = 24
    Top = 140
    Width = 422
    Height = 40
    Caption = '2. Testar Excel/OLE (Captura HWND -> Mata apenas o PID)'
    TabOrder = 2
    OnClick = btnExcelPIDClick
  end
  object btnFecharNotepad: TButton
    Left = 24
    Top = 190
    Width = 422
    Height = 40
    Caption = '3. Finalizar por Nome (.EXE) '#8221' Fechar todos notepad.exe'
    TabOrder = 3
    OnClick = btnFecharNotepadClick
  end
end
