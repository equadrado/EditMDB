object Form1: TForm1
  Left = 99
  Top = 112
  Width = 552
  Height = 326
  Caption = 'Editor de Banco de Dados Access'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 13
    Top = 45
    Width = 31
    Height = 13
    Caption = 'Senha'
  end
  object Button1: TButton
    Left = 8
    Top = 8
    Width = 75
    Height = 25
    Caption = 'Arquivo'
    TabOrder = 0
    OnClick = Button1Click
  end
  object Edit1: TEdit
    Left = 96
    Top = 8
    Width = 193
    Height = 21
    Enabled = False
    TabOrder = 1
  end
  object bbRepara: TButton
    Left = 8
    Top = 104
    Width = 75
    Height = 25
    Caption = 'Reparar'
    Enabled = False
    TabOrder = 2
    OnClick = bbReparaClick
  end
  object bbCompacta: TButton
    Left = 8
    Top = 136
    Width = 75
    Height = 25
    Caption = 'Compactar'
    Enabled = False
    TabOrder = 3
    OnClick = bbCompactaClick
  end
  object bbConect: TButton
    Left = 8
    Top = 72
    Width = 75
    Height = 25
    Caption = 'Conectar'
    Enabled = False
    TabOrder = 4
    OnClick = bbConectClick
  end
  object bbDisponivel: TButton
    Left = 8
    Top = 200
    Width = 75
    Height = 25
    Enabled = False
    TabOrder = 5
    Visible = False
  end
  object bbDados: TButton
    Left = 8
    Top = 168
    Width = 75
    Height = 25
    Caption = 'Dados'
    Enabled = False
    TabOrder = 6
    OnClick = bbDadosClick
  end
  object bbFechar: TButton
    Left = 9
    Top = 254
    Width = 75
    Height = 25
    Caption = 'Fechar'
    TabOrder = 7
    OnClick = bbFecharClick
  end
  object edSenha: TEdit
    Left = 96
    Top = 40
    Width = 81
    Height = 21
    PasswordChar = '*'
    TabOrder = 8
  end
  object Panel1: TPanel
    Left = 352
    Top = 72
    Width = 185
    Height = 209
    TabOrder = 9
    object Label3: TLabel
      Left = 10
      Top = 7
      Width = 38
      Height = 13
      Caption = 'Campos'
    end
    object Label4: TLabel
      Left = 16
      Top = 184
      Width = 21
      Height = 13
      Caption = 'Tipo'
    end
    object sbOKCp: TSpeedButton
      Left = 158
      Top = 4
      Width = 23
      Height = 22
      Glyph.Data = {
        36030000424D3603000000000000360000002800000010000000100000000100
        18000000000000030000C40E0000C40E00000000000000000000C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        000000000000000000000000C0C0C0C0C0C0000000000000000000000000C0C0
        C0000000000000C0C0C0C0C0C0000000000000C0C0C0C0C0C0000000000000C0
        C0C0C0C0C0000000000000C0C0C0000000000000C0C0C0C0C0C0C0C0C0000000
        000000C0C0C0C0C0C0000000000000C0C0C0C0C0C00000000000000000000000
        00C0C0C0C0C0C0C0C0C0C0C0C0000000000000C0C0C0C0C0C0000000000000C0
        C0C0C0C0C0000000000000000000C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0000000
        000000C0C0C0C0C0C0000000000000C0C0C0C0C0C0000000000000C0C0C00000
        00C0C0C0C0C0C0C0C0C0C0C0C0000000000000C0C0C0C0C0C0000000000000C0
        C0C0C0C0C0000000000000C0C0C0C0C0C0000000C0C0C0C0C0C0C0C0C0C0C0C0
        000000000000000000000000C0C0C0C0C0C0000000000000000000C0C0C00000
        00000000000000C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0}
      Visible = False
      OnClick = sbOKCpClick
    end
    object ListBox2: TListBox
      Left = 10
      Top = 30
      Width = 105
      Height = 146
      ItemHeight = 13
      TabOrder = 0
      OnClick = ListBox2Click
    end
    object bbNovoCp: TButton
      Left = 125
      Top = 31
      Width = 51
      Height = 18
      Caption = 'Novo'
      Enabled = False
      TabOrder = 1
      OnClick = bbNovoCpClick
    end
    object bbExclCp: TButton
      Left = 125
      Top = 55
      Width = 51
      Height = 18
      Caption = 'Excluir'
      Enabled = False
      TabOrder = 2
      OnClick = bbExclCpClick
    end
    object Button4: TButton
      Left = 125
      Top = 79
      Width = 51
      Height = 18
      Caption = 'Alterar'
      Enabled = False
      TabOrder = 3
    end
    object Button7: TButton
      Left = 125
      Top = 103
      Width = 51
      Height = 18
      Caption = 'Rela'#231#227'o'
      Enabled = False
      TabOrder = 4
    end
    object ComboBox1: TComboBox
      Left = 48
      Top = 181
      Width = 130
      Height = 21
      ItemHeight = 13
      TabOrder = 5
      Items.Strings = (
        'TEXT(10)'
        'LONGBINARY'
        'CURRENCY'
        'CURRENCY NOT NULL'
        'YESNO'
        'AUTOINCREMENT'
        'BYTE'
        'BYTE NOT NULL'
        'LONG'
        'LONG NOT NULL'
        'DATETIME'
        'DATETIME NOT NULL'
        'DATE'
        'DATE NOT NULL'
        '')
    end
    object edNomeCp: TEdit
      Left = 56
      Top = 4
      Width = 97
      Height = 21
      TabOrder = 6
      Visible = False
      OnKeyPress = edNomeTbKeyPress
    end
  end
  object Panel2: TPanel
    Left = 96
    Top = 72
    Width = 241
    Height = 209
    TabOrder = 10
    object Label2: TLabel
      Left = 8
      Top = 8
      Width = 38
      Height = 13
      Caption = 'Tabelas'
    end
    object sbOKTb: TSpeedButton
      Left = 214
      Top = 5
      Width = 23
      Height = 22
      Glyph.Data = {
        36030000424D3603000000000000360000002800000010000000100000000100
        18000000000000030000C40E0000C40E00000000000000000000C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        000000000000000000000000C0C0C0C0C0C0000000000000000000000000C0C0
        C0000000000000C0C0C0C0C0C0000000000000C0C0C0C0C0C0000000000000C0
        C0C0C0C0C0000000000000C0C0C0000000000000C0C0C0C0C0C0C0C0C0000000
        000000C0C0C0C0C0C0000000000000C0C0C0C0C0C00000000000000000000000
        00C0C0C0C0C0C0C0C0C0C0C0C0000000000000C0C0C0C0C0C0000000000000C0
        C0C0C0C0C0000000000000000000C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0000000
        000000C0C0C0C0C0C0000000000000C0C0C0C0C0C0000000000000C0C0C00000
        00C0C0C0C0C0C0C0C0C0C0C0C0000000000000C0C0C0C0C0C0000000000000C0
        C0C0C0C0C0000000000000C0C0C0C0C0C0000000C0C0C0C0C0C0C0C0C0C0C0C0
        000000000000000000000000C0C0C0C0C0C0000000000000000000C0C0C00000
        00000000000000C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0
        C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0}
      Visible = False
      OnClick = sbOKTbClick
    end
    object ListBox1: TListBox
      Left = 8
      Top = 29
      Width = 169
      Height = 169
      ItemHeight = 13
      TabOrder = 0
      OnClick = ListBox1Click
    end
    object bbNovoTb: TButton
      Left = 181
      Top = 31
      Width = 51
      Height = 18
      Caption = 'Novo'
      Enabled = False
      TabOrder = 1
      OnClick = bbNovoTbClick
    end
    object bbExclTb: TButton
      Left = 181
      Top = 55
      Width = 51
      Height = 18
      Caption = 'Excluir'
      Enabled = False
      TabOrder = 2
      OnClick = bbExclTbClick
    end
    object edNomeTb: TEdit
      Left = 56
      Top = 4
      Width = 153
      Height = 21
      TabOrder = 3
      Visible = False
      OnKeyPress = edNomeTbKeyPress
    end
  end
  object cbTiraSenha: TCheckBox
    Left = 200
    Top = 40
    Width = 97
    Height = 17
    Caption = 'Tirar senha'
    TabOrder = 11
  end
  object RadioGroup1: TRadioGroup
    Left = 296
    Top = 6
    Width = 105
    Height = 57
    Caption = 'Vers'#227'o Ini do BD'
    ItemIndex = 0
    Items.Strings = (
      'Access 97'
      'Access 2000')
    TabOrder = 12
  end
  object RadioGroup2: TRadioGroup
    Left = 416
    Top = 6
    Width = 105
    Height = 57
    Caption = 'Vers'#227'o Fim do BD'
    ItemIndex = 0
    Items.Strings = (
      'Access 97'
      'Access 2000')
    TabOrder = 13
  end
  object OpenDialog1: TOpenDialog
    DefaultExt = '*.mdb'
    Left = 256
    Top = 8
  end
  object ADOConnection1: TADOConnection
    LoginPrompt = False
    Left = 16
    Top = 8
  end
  object DataSource1: TDataSource
    DataSet = ADODataSet1
    Left = 176
    Top = 8
  end
  object ADOCommand1: TADOCommand
    Connection = ADOConnection1
    Parameters = <>
    Left = 136
    Top = 8
  end
  object ADODataSet1: TADODataSet
    Connection = ADOConnection1
    Parameters = <>
    Left = 208
    Top = 8
  end
end
