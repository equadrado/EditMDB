unit UEditaMDB;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, StdCtrls,
  OleServer, JRO_TLB, Comobj, ExtCtrls, Buttons {para reparar Bando de dados};

type
  TForm1 = class(TForm)
    OpenDialog1: TOpenDialog;
    Button1: TButton;
    bbRepara: TButton;
    ADOConnection1: TADOConnection;
    bbCompacta: TButton;
    bbConect: TButton;
    bbDisponivel: TButton;
    bbDados: TButton;
    bbFechar: TButton;
    edSenha: TEdit;
    Label1: TLabel;
    DataSource1: TDataSource;
    ADOCommand1: TADOCommand;
    Edit1: TEdit;
    ADODataSet1: TADODataSet;
    Panel1: TPanel;
    Label3: TLabel;
    ListBox2: TListBox;
    bbNovoCp: TButton;
    bbExclCp: TButton;
    Button4: TButton;
    Panel2: TPanel;
    Label2: TLabel;
    ListBox1: TListBox;
    bbNovoTb: TButton;
    bbExclTb: TButton;
    Button7: TButton;
    Label4: TLabel;
    ComboBox1: TComboBox;
    edNomeTb: TEdit;
    edNomeCp: TEdit;
    sbOKTb: TSpeedButton;
    sbOKCp: TSpeedButton;
    cbTiraSenha: TCheckBox;
    RadioGroup1: TRadioGroup;
    RadioGroup2: TRadioGroup;
    function Repair(sNewMDB : String) : Boolean;
    procedure Button1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bbReparaClick(Sender: TObject);
    procedure bbFecharClick(Sender: TObject);
    procedure bbConectClick(Sender: TObject);
    procedure bbCompactaClick(Sender: TObject);
    procedure ListBox1Click(Sender: TObject);
    procedure bbDadosClick(Sender: TObject);
    procedure bbExclCpClick(Sender: TObject);
    procedure ListBox2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure bbExclTbClick(Sender: TObject);
    procedure edNomeTbKeyPress(Sender: TObject; var Key: Char);
    procedure bbNovoTbClick(Sender: TObject);
    procedure bbNovoCpClick(Sender: TObject);
    procedure sbOKTbClick(Sender: TObject);
    procedure sbOKCpClick(Sender: TObject);
    function CompactaBD(ArqOrigem, SenOrigem, ArqDestino, SenDestino : string) : boolean;
  private
    { Private declarations }
  public
    { Public declarations }
    cDirBD, cNomeArq, cNomeTabela, cNomeCampo: string;
  end;

var
  Form1: TForm1;
const
   SProvider4 = 'Provider=Microsoft.Jet.OLEDB.4.0;'+
               'Jet OLEDB:Engine Type=4;'+   // Access 97
               'Data Source=';
   SProvider5 = 'Provider=Microsoft.Jet.OLEDB.4.0;'+
               'Jet OLEDB:Engine Type=5;'+   // Access 97
               'Data Source=';
   SProvider6 = 'Provider=Microsoft.Jet.OLEDB.4.0;'+
               'Jet OLEDB:Engine Type=6;'+   // Access 97
               'Data Source=';

implementation

uses UFuncDivs, UDados;

{$R *.dfm}

function TForm1.Repair(sNewMDB : String) : Boolean;
var
   DBEngine: Variant;
   nBD : integer;
   cNomeBD, cNomeBDTemp, cNomeBDParc : String;
begin
   Result := False;

   cNomeBD := cDirBD+cNomeArq;
   cNomeBDTemp := cDirBD+'Cpkt'+sNewMDB;

   if Application.MessageBox (PChar('O Banco de Dados '+sNewMDB+' será reparado !'),
         'Atenção',MB_YESNO+MB_ICONQUESTION) = idYes then
   begin

      if FileExists(cNomeBDTemp) then
         DeleteFile(cNomeBDTemp);

      try
         if CopyFile(PChar(cNomeBD), PChar(cNomeBDTemp), False) then
         begin
            Sleep(5000);
            { Get a handle  to the Database  Engine. }
            DBEngine  := CreateOleObject('DAO.DBEngine.35');
            { Repair the  database.  }
            DBEngine.RepairDatabase(cNomeBD);

            Application.MessageBox (PChar('Banco de Dados reparado !'#13#10+
                  'Arquivo com erro salvo como: '+cNomeBDTemp+' !'),
                  'Concluído',MB_OK+MB_ICONEXCLAMATION);

            Result := True;
         end;
      except
         on E: EOleException do
         begin

            Application.MessageBox (PChar('Erro na reparação do Banco de Dados:'#13#10+
                  sNewMDB+'.mdb. Tente novamente !'+#13#10+
                  E.Message),
                  'Erro',MB_OK+MB_ICONERROR);
         end;
      end;  // try
   end;
end;


procedure TForm1.Button1Click(Sender: TObject);
begin
   if OpenDialog1.Execute then
   begin
      Edit1.Text := OpenDialog1.FileName;
      cDirBD := ExtractFilePath(Edit1.Text);
      cNomeArq := ExtractFileName(Edit1.Text);

      bbConect.Enabled := True;
      bbRepara.Enabled := True;
      bbCompacta.Enabled := True;

      bbConect.SetFocus;
   end;
end;

procedure TForm1.FormShow(Sender: TObject);
begin
   OpenDialog1.InitialDir := ExtractFilePath(Application.ExeName);
end;

procedure TForm1.bbReparaClick(Sender: TObject);
begin
   ReparaMDB(Edit1.Text, False);
end;

procedure TForm1.bbFecharClick(Sender: TObject);
begin
   Form1.Close;
end;

procedure TForm1.bbConectClick(Sender: TObject);
begin
   try
      ADOConnection1.Connected := False;

      if bbConect.Caption = 'Desconectar' then
      begin
         bbConect.Caption := 'Conectar';
         bbDados.Enabled := False;
         bbNovoTb.Enabled := False;
         bbExclTb.Enabled := False;
         bbNovoCp.Enabled := False;
         bbExclCp.Enabled := False;
         bbCompacta.Enabled := True;
         bbRepara.Enabled := True;
         ListBox1.Items.Clear;
         ListBox2.Items.Clear;
         Exit;
      end;

      ADOConnection1.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;'+
              'Password="";'+
              'User ID=Admin;'+
              'Data Source='+Edit1.Text+';'+
              'Mode=Share Deny None;'+
              'Extended Properties="";'+
              'Locale Identifier=1033;'+
              'Jet OLEDB:System database="";'+
              'Jet OLEDB:Registry Path="";'+
              'Jet OLEDB:Database Password="'+edSenha.Text+'";'+
              'Jet OLEDB:Engine Type=4;'+  // Access 97
              'Jet OLEDB:Database Locking Mode=0;'+
              'Jet OLEDB:Global Partial Bulk Ops=2;'+
              'Jet OLEDB:Global Bulk Transactions=1;'+
              'Jet OLEDB:New Database Password="";'+
              'Jet OLEDB:Create System Database=False;'+
              'Jet OLEDB:Encrypt Database=False;'+
              'Jet OLEDB:Compact Without Replica Repair=False;'+
              'Jet OLEDB:SFP=False;';

      ADOConnection1.Connected := True;
      bbConect.Caption := 'Desconectar';
      bbCompacta.Enabled := False;
      bbRepara.Enabled := False;
      bbNovoTb.Enabled := True;

      ListBox1.Items.Clear;
      ADOConnection1.GetTableNames(ListBox1.Items, False);
      bbDados.Enabled := False;
   except
   end;
end;

procedure TForm1.bbCompactaClick(Sender: TObject);
var
   cSen : string;
begin
   if cbTiraSenha.Checked then
      cSen := ''
   else
      cSen := edSenha.Text;
   CompactaBD(cDirBD+cNomeArq, edSenha.Text, cDirBD+'Cpkt'+cNomeArq, cSen);
end;

procedure TForm1.ListBox1Click(Sender: TObject);
var
   a : integer;
begin
   for a := 0 to ListBox1.Items.Count-1 do
   begin
      if ListBox1.Selected[a] then
      begin
         ListBox2.Items.Clear;
         cNomeTabela := ListBox1.Items.Strings[a];
         ADOConnection1.GetFieldNames(cNomeTabela, ListBox2.Items);
         bbDados.Enabled := True;
         bbExclTb.Enabled := True;
         bbNovoCp.Enabled := True;
      end
   end;
end;

procedure TForm1.bbDadosClick(Sender: TObject);
begin
   ADODataSet1.Close;
   ADODataSet1.CommandText := 'SELECT * FROM '+cNomeTabela;
   ADODataSet1.Open;

   fmDados.ShowModal;

   ADODataSet1.Close;
end;

procedure TForm1.bbExclCpClick(Sender: TObject);
begin
   ADOCommand1.CommandText := 'ALTER TABLE '+cNomeTabela+
                              ' DROP COLUMN '+cNomeCampo;
   ADOCommand1.Execute;

   ListBox2.Items.Clear;
   ADOConnection1.GetFieldNames(cNomeTabela, ListBox2.Items);
end;

procedure TForm1.ListBox2Click(Sender: TObject);
var
   a : integer;
begin
   for a := 0 to ListBox2.Items.Count-1 do
   begin
      if ListBox2.Selected[a] then
         cNomeCampo := ListBox2.Items.Strings[a];
   end;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
   ADOConnection1.Connected := False;
end;

procedure TForm1.bbExclTbClick(Sender: TObject);
begin
   ADOCommand1.CommandText := 'DROP TABLE '+cNomeTabela;
   ADOCommand1.Execute;

   ListBox2.Items.Clear;
   ListBox1.Items.Clear;
   ADOConnection1.GetTableNames(ListBox1.Items, False);
   bbDados.Enabled := False;
   bbExclTb.Enabled := False;
end;

procedure TForm1.edNomeTbKeyPress(Sender: TObject; var Key: Char);
begin
   if key = #27 then
   begin
      ListBox1.SetFocus;
      edNomeTb.Visible := False;
      edNomeCp.Visible := False;
      sbOKTb.Visible := False;
      sbOKCp.Visible := False;
   end
   else if Pos(key,'çÇãÃõÕúÚíÍáÁóÓéÉüÜ ,.;:~^/\|?[{]}=+-)("!@#$%¨&*') > 0 then
      Key := #0;
end;

procedure TForm1.bbNovoTbClick(Sender: TObject);
begin
   edNomeTb.Visible := True;
   sbOKTb.Visible := True;
   edNomeTb.SetFocus;
end;

procedure TForm1.bbNovoCpClick(Sender: TObject);
begin
   edNomeCp.Visible := True;
   sbOKCp.Visible := True;
   edNomeCp.SetFocus;
end;

procedure TForm1.sbOKTbClick(Sender: TObject);
begin
   try
      ADOCommand1.CommandText := 'CREATE TABLE '+edNomeTb.Text;
      ADOCommand1.Execute;
   except
   end;

   ListBox2.Items.Clear;
   ListBox1.Items.Clear;
   ADOConnection1.GetTableNames(ListBox1.Items, False);

   edNomeTb.Visible := False;
   sbOKTb.Visible := False;
   ListBox1.SetFocus;
end;

procedure TForm1.sbOKCpClick(Sender: TObject);
begin
   try
      ADOCommand1.CommandText := 'ALTER TABLE '+cNomeTabela+
                                 ' ADD COLUMN '+edNomeCp.Text+
                                 ' '+ComboBox1.Text;
      ADOCommand1.Execute;
   except
   end;

   ListBox2.Items.Clear;
   ADOConnection1.GetFieldNames(cNomeTabela, ListBox2.Items);

   edNomeCp.Visible := False;
   sbOKCp.Visible := False;
   ListBox2.SetFocus;
end;

function TForm1.CompactaBD(ArqOrigem, SenOrigem, ArqDestino, SenDestino : string) : boolean;
var
   nBD : integer;

   JE : TJetEngine; //Jet Engine
   cMsg : String;
begin
   try
      Result := False;
      cMsg := 'Falha na compactação !';
      if FileExists(ArqDestino) then
         DeleteFile(ArqDestino);

      if FileExists(ArqOrigem) and
         not FileExists(BuscaTroca(UpperCase(ArqOrigem),'.MDB','.LDB')) then
      begin
         JE:= TJetEngine.Create(Application);
         try
            JE.CompactDatabase(
                               IIf(RadioGroup1.ItemIndex = 0,
                                    SProvider4, SProvider5)+
                               ArqOrigem+
               ';Mode=Share Exclusive;'+
               'Jet OLEDB:Database Password='+SenOrigem,
                               IIf(RadioGroup2.ItemIndex = 0,
                                    SProvider4, SProvider5)+
                               ArqDestino+
               ';Mode=Share Exclusive;'+
               'Jet OLEDB:Database Password='+SenDestino);

            cMsg := 'Compactação concluída';
            Result := True;
         except
            on E:Exception do
            begin
               Result := False;
               JE.FreeOnRelease;
               cMsg :=  E.Message+': '+ArqOrigem;
            end;
         end;
      end;
   finally
      ShowMessage(cMsg);
   end;
end;

end.
