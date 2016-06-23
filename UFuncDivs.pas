unit UFuncDivs;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, ExtCtrls, Buttons, ComCtrls, StdCtrls, JPEG, Math,
  DB, DbTables, OleServer, ComObj, JRO_TLB, Registry, QRPRNTR,
  Printers, DateUtils, StrUtils, ShellAPI, ShlObj, ActiveX,
  xmldom, XMLIntf, XMLDoc;


type
   TExec = procedure of object;

   procedure LimpaTeclado;
   function NumToBol(num: Integer): Boolean;
   function BolToNum(bol: Boolean): Integer;
   function StrToBol(str: string): Boolean;
   function BolToStr(bol: Boolean): string;
   function BolToStrP(bol: Boolean): string;
   function IIfStr( bCond: Boolean; cVar1: string; cVar2: string): string;
   function IIf( bCond: Boolean; cVar1: variant; cVar2: variant): variant;
   function AdSeqChar(cSeqChar: string; nChar: integer): string;
   function InverteData(dData: TDate): string;
   function TestaData(Data: String):Boolean;
   function Before(const Search, Find: widestring): widestring;
   function after(const Search, Find: widestring): widestring;
   function IntPrime(Value: Integer): Boolean;
   function PosMid(cSubstr, cString : string; nPosIni : integer) : integer;
   function RPosMid(cSubstr, cString : string; nPosFim : integer) : integer;
   function BuscaTroca(Text,Busca,Troca : string) : string;
   function StrAntes(cString: string ; cAntes: Char; nTamanho: integer) : string;
   function StrDepois(cString: string ; cAntes: Char; nTamanho: integer) : string;
   function StrZero(cString: string ; nTamanho: integer) : string;
   function IntToStrZero(nInteiro: integer ; nTamanho: integer) : string;
   function IdadeN(Nascimento:TDateTime; lCompleto : boolean; lSoNum : boolean = False; Hoje:TDateTime = 0) : String;
   function PriMaiuscula(Texto:String): String;
   function LimpaAcentos(Texto:String): String;
   function LimpaNome(Texto:String): String;
   function Replicate( Caracter:String; Quant:Integer ): String;
   function Arredonda(nNumero: Real; nCasaDec: integer): string;
   function Arredonda2(nNumero: Extended; nCasaDec: Integer): Extended;
   function exp2(X: Single): Single;
   function DiasUteis( DataIni:TDate; nDias: integer ): TDate;
   function IdGestC( DataIni:TDateTime; DataFim:TDateTime ): String;
   function IdGestN( DataIni:TDateTime; DataFim:TDateTime ): single;
   function DataIGCDat(cIgest: String; DataIni: TDateTime): string;
   function IdGestNtoC( nIdGest: single ): String;
   function IdGestCtoN( cIdGest: string ): integer;
   function BuscaArquivos(cDiretorio: string; cExt: string ): TStringList;
   function BuscaVetorStr(StrVetor: array of string; SubStr: string ): integer;
   function BuscaVetorInt(StrVetor: array of integer; SubStr: integer): integer;
   function BuscaVetorReal(var StrVetor: array of Real; SubStr: Real): Real;
   function BiosDate: String;
   function GetBuildInfo:string;
   function WinExecAndWait32( FileName: String; Visibility  : Integer ) : Cardinal;
   procedure PostKeyEx32(Key: Word; const Shift: TShiftState; SpecialKey: boolean);
   function AlteraDataHora (dDataHora : TDateTime) : boolean ;
   function DriveSpace(Drive: string): integer;
   function DriveOk(Drive: string): boolean;
   procedure AddHexString(S : String; Lines : TStrings );
   function CaptureScreenRect( ARect: TRect ): TJPEGImage;
   function CompactaBD(ArqOrigem, SenOrigem, ArqDestino, SenDestino : string; lInfo : boolean) : boolean;
   procedure SetDefaultPrinter(PrinterName: String);
   function ReparaMDB(cNomeMDB : String; lAutom : boolean) : Boolean;
   function Maiuscula(Texto:String): String;
   function PrimaPalavra(Texto:String; lLimpa: boolean): String;
   function Mes(dData : TDate) : integer;
   function Ano(dData : TDate) : integer;
   function DateTimeToSQLDateTimeString(Data: TDateTime; Format: string;
                          OnlyDate: Boolean = True): string;
   function ValidaCPF(Const s: ShortString): Boolean;
   function ValidaCNPJ(const s: string): boolean;
   function ValidaEMail2(const EMailIn: string):Boolean;
   function ValidaEmail(email: string): boolean;
   function QuotedData(dData : TDateTime; nTpBD : integer) : string;
   function QuotedDataHora(dData : TDateTime; nTpBD : integer; cFormato : string = 'dd/mm/yyyy hh:nn:ss') : string;
   function QuotedTexto(cTexto : string; nTpBD : integer) : string;
   function QuotedBool(lBoll : boolean; nTpBD : integer) : string;
   function QuotedMoney(cValor : string; nTpBD : integer) : string;
   function NomeMaquina: string;
   function IdWin: string;
   function GetCommonPathFiles : String;
   function ListarArquivos(Diretorio: string; Sub:Boolean) : TStrings;
   function TemAtributo(Attr, Val: Integer): Boolean;
   function DelTree(DirName : string): Boolean;
   procedure DelTree2(const Directory: TFileName);
   function FindFile(const filespec: TFileName; attributes: integer): TStringList;
   function ContaStr(Text, Busca : string) : integer;
   procedure CriarAtalhoIni(FileName, Parameters, InitialDir,
                                    ShortcutName, ShortcutFolder : String);
   procedure CriarAtalhoDT(FileName, Parameters, InitialDir,
                                 ShortcutName, ShortcutFolder : String);
   function TirarAcentos( cString: String ): String;
   function TamanhoPapel(cTamPapel : string) : TQRPaperSize;
   function ChecaImpr : boolean;
   function IndexImpr(cImpr : string) : integer;
   function PesqTagNode(const Strtag: String; NodeIni : IXmlNode): IXmlNode;
   function CopyFiles(FromPath, ToPath, FileMask: String): Boolean;

implementation

procedure LimpaTeclado;
var
  msg: TMsg;
begin
  while PeekMessage(msg, 0, WM_KEYFIRST, WM_KEYLAST, PM_REMOVE or PM_NOYIELD) do;
end;

function NumToBol(num: Integer): Boolean;
begin
   if num = 1 then
      Result := true
   else
      Result := false;
end;

function BolToNum(bol: Boolean): Integer;
begin
   if bol = true then
      Result := 1
   else
      Result := 0;
end;

function StrToBol(str: string): Boolean;
begin
   if (Pos('u', str) > 0) or (Pos('U', str) > 0) then
      Result := True
   else
      Result := False;
end;

function BolToStr(bol: Boolean): string;
begin
   if bol then
      Result := 'True'
   else
      Result := 'False';
end;

function BolToStrP(bol: Boolean): string;
begin
   if bol then
      Result := 'Sim'
   else
      Result := 'N„o';
end;

function IIfStr( bCond: Boolean; cVar1: string; cVar2: string): string;
begin
   if bCond = true then
      Result := cVar1
   else
      Result := cVar2;
end;

function IIf( bCond: Boolean; cVar1: variant; cVar2: variant): variant;
begin
   if bCond = True then
      Result := cVar1
   else
      Result := cVar2;
end;

function AdSeqChar(cSeqChar: string; nChar: integer): string;
var
   aCodLet: array of integer;
   cSeqLetras: string;
   a: integer;
   nCharSeq: integer;
   lErroSeq: boolean;

begin
   nCharSeq := Length(cSeqChar);
   if nChar < 2 then
      nChar := 2
   else if nChar < nCharSeq then
      nChar := nCharSeq
   else if nChar > nCharSeq then
      begin
         cSeqChar := StringOfChar('0', nChar-nCharSeq)+cSeqChar;
      end;

   if nCharSeq = 0 then
      begin
         Result := StringOfChar('0', nChar);
      end
   else
      begin
         cSeqLetras := '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
         SetLength(aCodLet, nChar+1);
         lErroSeq := false;

         for a := 1 to nChar do
         begin
            aCodLet[a] := Pos( Copy(cSeqChar,a,1), cSeqLetras );
            if aCodLet[a] = 0 then
               lErroSeq := true;
         end;

         if not lErroSeq then
         begin
            for a := nChar downto 2 do
            begin
               aCodLet[a] := aCodLet[a] + 1;
               if aCodLet[a] > 36 then
                  begin
                     aCodLet[a] := 1;
                     if aCodLet[a-1] < 36 then
                        begin
                           aCodLet[a-1] := aCodLet[a-1] + 1;
                           Break;
                        end;
                  end
               else
                  Break;
            end;
            if aCodLet[1] > 36 then
               lErroSeq := true;
         end;
         if lErroSeq then
            begin
               Application.MessageBox('ERRO NA SEQ‹ NCIA DE C”DIGOS !',
                                  'Erro', mb_OK+mb_IconAsterisk);
               Result := StringOfChar('?',nChar);
            end
         else
            begin
               Result := '';
               for a := 1 to nChar do
                  Result := Result + Copy(cSeqLetras, aCodLet[a], 1);
            end;
      end;
end;

function InverteData(dData: TDate): string;
var
   nAno, nMes, nDia : word;
begin
   Result := '00/00/0000';
   DecodeDate(dData, nAno, nMes, nDia);
   Result := StrAntes(IntToStr(nMes),'0',2)+'/'+
             StrAntes(IntToStr(nDia),'0',2)+'/'+
             StrAntes(IntToStr(nAno),'0',4);
end;

function TestaData(Data: String):Boolean;
{Testa se uma data È v·lida ou n„o}
var
   nPos, Dia, Mes, Ano: integer;
   DataTeste : string;
const
   aDiaMesNl : array[1..12] of integer = (31,28,31,30,31,30,31,31,30,31,30,31);
   aDiaMesBi : array[1..12] of integer = (31,29,31,30,31,30,31,31,30,31,30,31);
begin
   Result := false;
   DataTeste := Data;
   if BuscaTroca(DataTeste,' ','') <> '//' then
   begin
      nPos := Pos('/',Data);
      if nPos <> 0 then
      begin
         Dia := StrToInt(Copy(Data,1,nPos-1));
         Data := Copy(Data,nPos+1,8);
         nPos := Pos('/',Data);
         if nPos <> 0 then
         begin
            Mes := StrToInt(Copy(Data,1,nPos-1));
            if (Mes > 0) and (Mes < 13) then
            begin
               Ano := StrToInt(Copy(Data,nPos+1,4));
               if (Ano < 1850) or (Ano > 2100) then
                  Result := false
               else if (Ano mod 4 <> 0) and (Dia > 0) then
               begin
                  if (Dia <= aDiaMesNl[Mes]) then
                     Result := true
               end
               else if (Dia <= aDiaMesBi[Mes]) then
                  Result := true
            end;
         end;
      end;
   end;
end;

function Before(const Search, Find: widestring): widestring;
{Retorna uma cadeia de caracteres antecedentes
 a uma parte da string selecionada}
const
   BlackSpace = [#33..#126];
var
   index: integer;
begin
   index:=Pos(Find, Search);
   if index = 0 then
      Result:=Search
   else
      Result:=Copy(Search, 1, index - 1);
end;

function after(const Search, Find: widestring): widestring;
{Retorna uma cadeia de caracteres apÛs a parte
 da string selecionada}
var
   index: integer;
begin
   index := Pos(Find, Search);
   if index = 0 then
      begin
         Result := '';
      end
   else
      begin
         Result := Copy(Search, index + Length(Find), Length(Search));
      end;
end;

function IntPrime(Value: Integer): Boolean;
{Testa se um numero È primo ou n„o}
var
i : integer;
begin
Result := False;
Value := Abs(Value);
if Value mod 2 <> 0 then
   begin
   i := 1;
   repeat
   i := i + 2;
   Result:= Value mod i = 0
   until Result or ( i > Trunc(sqrt(Value)) );
         Result:= not Result;
   end;
end;

function PosMid(cSubstr, cString : string; nPosIni : integer) : integer;
var
   cSubTex : string;
begin
   cSubTex := Copy(cString, nPosIni, Length(cString)-nPosIni+1);

   Result := Pos(cSubstr, cSubTex);

   if Result > 0 then
      Result := Result + nPosIni - 1;
end;

function RPosMid(cSubstr, cString : string; nPosFim : integer) : integer;
var
   cSubTex : string;
   nPosF, nPosR : integer;
begin
   cSubTex := Copy(cString, 1, nPosFim);

   nPosR := Pos(cSubstr, cSubTex);
   nPosF := nPosR;

   while nPosF > 0 do
   begin
      nPosR := nPosF;
      nPosF := PosMid(cSubstr, cSubTex, nPosF+1);
   end;

   Result := nPosR;
end;

function BuscaTroca(Text, Busca, Troca : string) : string;
{ Substitui um caractere dentro da string}
var
   nBusca, nPos : integer;
begin
   nBusca := Length(Busca);
   nPos := Pos(Busca, Text);

   while (nPos <> 0) and (Busca <> Troca) do
   begin
      Delete(Text, nPos, nBusca);
      Insert(Troca, Text, nPos);
      nPos := PosMid(Busca, Text, nPos+Length(Troca));
   end;
   Result := Text;
end;

function StrAntes(cString: string; cAntes: Char; nTamanho: integer) : string;
begin
   nTamanho := nTamanho - Length(cString);
   if nTamanho > 0 then
      Result := StringOfChar(cAntes, nTamanho)+cString
   else
      Result := cString;
end;

function StrDepois(cString: string; cAntes: Char; nTamanho: integer) : string;
begin
   nTamanho := nTamanho - Length(cString);
   if nTamanho > 0 then
      Result := cString+StringOfChar(cAntes, nTamanho)
   else
      Result := cString;
end;

function StrZero(cString: string ; nTamanho: integer) : string;
begin
   nTamanho := nTamanho - Length(cString);
   if nTamanho > 0 then
      Result := StringOfChar('0', nTamanho)+cString
   else
      Result := cString;
end;

function IntToStrZero(nInteiro: integer ; nTamanho: integer) : string;
var
   cString : string;
begin
   cString := IntToStr(nInteiro);

   Result := StrZero(cString, nTamanho);
end;

function IdadeN(Nascimento:TDateTime; lCompleto : boolean; lSoNum : boolean = False; Hoje:TDateTime = 0) : String;
//
// Retorna a idade de uma pessoa a partir da data de nascimento
//
   type
      TpData = Record
             Ano : Word;
             Mes : Word;
             Dia : Word;
   end;
const
   Qdm:String = '312831303130313130313031';          // Qtde dia no mes
var
   Dth : TpData;                                  // Data de hoje
   Dtn : TpData;                                  // Data de nascimento
   anos, meses, dias, nrd : Shortint;           // Usadas para calculo da idade
begin
   if Hoje = 0 then
      Hoje := Date;
   DecodeDate(Hoje,Dth.Ano,Dth.Mes,Dth.Dia);
   DecodeDate(Nascimento,Dtn.Ano,Dtn.Mes,Dtn.Dia);
   anos := Dth.Ano - Dtn.Ano;
   meses := Dth.Mes - Dtn.Mes;
   dias := Dth.Dia - Dtn.Dia;
   if dias < 0 then
   begin
      nrd := StrToInt(Copy(Qdm,(Dth.Mes-1)*2-1,2));
      if ((Dth.Mes-1)=2) and ((Dth.Ano Div 4)=0) then
         Inc(nrd);

      dias := dias+nrd;
      Dec(meses); // tira 1 do mes
   end;
   if meses < 0 then
   begin
      Dec(anos);
      meses := meses+12;
   end;
   if meses = 12 then
   begin
      meses := 0;
      Inc(anos);
   end;
   if lCompleto then
   begin
      if lSoNum then
         Result := StrZero(IntToStr(anos), 3)+
                   StrZero(IntToStr(meses), 2)+
                   StrZero(IntToStr(dias), 2)
      else
         Result := IntToStr(anos)+'a'+IntToStr(meses)+'m'+IntToStr(dias)+'d'
   end
   else
   begin
      if lSoNum then
         Result := StrZero(IntToStr(anos), 3)+StrZero(IntToStr(meses), 2)
      else
         Result := IntToStr(anos)+'a'+IntToStr(meses)+'m';
   end;
end;

function PriMaiuscula(Texto:String): String;
{Converte a primeira letra do texto especificado para
 maiuscula e as restantes para minuscula}
var
   nPos,nFim: Integer;
   cParc : string;
begin
   Result := '';
   Texto := Trim(Texto);
   Texto := BuscaTroca(Texto, '.',' ');
   Texto := BuscaTroca(Texto, ',',' ');
   Texto := BuscaTroca(Texto, '√','„');
   Texto := BuscaTroca(Texto, '’','ı');
   Texto := BuscaTroca(Texto, '«','Á');
   Texto := BuscaTroca(Texto, '  ',' ');
   Texto := BuscaTroca(Texto, Chr(39), Chr(39)+' '); // '
   Texto := BuscaTroca(Texto, '  ',' ');
   if Texto <> '' then
   begin
      nPos := Length(Texto);
      while nPos > 0 do
      begin
         nPos := Pos(' ',Texto);
         if nPos = 0 then
            nFim := Length(Texto)
         else
            nFim := nPos;

         cParc := Copy(Texto, 1, nFim);

         if (Pos('*'+cParc,'*de *da *do *das *dos ') = 0) and
                                    (Pos('(', cParc) = 0) then
            Result := Result + UpperCase(Copy(cParc,1,1))+
                      LowerCase(Copy(cParc,2,nFim-1))
         else
            Result := Result + cParc;

         Texto := Copy(Texto,nFim+1,255);
      end;
   end;
end;

function LimpaAcentos(Texto:String): String;
const
 accent   : string = '„‡·‰‚ËÈÎÍÏÌÔÓıÚÛˆÙ˘˙¸˚Á√¿¡ƒ¬»…À ÃÕœŒ’“”÷‘Ÿ⁄‹€«Ò—.-/'+Chr(39); // '
 noaccent : string = 'aaaaaeeeeiiiiooooouuuucAAAAAEEEEIIIIOOOOOUUUUCnN    ';
var
   a : integer;
begin
   Result := Texto;
   for a := 0 to Length(accent)-1 do
      Result := BuscaTroca(Result, accent[a], noaccent[a]);
end;

function LimpaNome(Texto:String): String;
const
 accent   : string = '‡·‰‚ËÈÎÍÏÌÔÓÚÛˆÙ˘˙¸˚¿¡ƒ¬»…À ÃÕœŒ“”÷‘Ÿ⁄‹€Ò—.-/'+Chr(39); // '
 noaccent : string = 'aaaaeeeeiiiioooouuuuAAAAEEEEIIIIOOOOUUUUnN    ';
var
   a : integer;
begin
   Result := Texto;
   for a := 0 to Length(accent)-1 do
      Result := BuscaTroca(Result, accent[a], noaccent[a]);

end;

function Replicate( Caracter:String; Quant:Integer ): String;
{Repete o mesmo caractere v·rias vezes}
var I : Integer;
begin
  Result := '';
  for I := 1 to Quant do
      Result := Result + Caracter;
end;

function Arredonda(nNumero: Real; nCasaDec: integer): string;
const
   aNumCasa : array [0..10] of string = ('0', '0.0', '0.00',
      '0.000', '0.0000', '0.00000', '0.000000', '0.0000000',
      '0.00000000', '0.000000000', '0.0000000000');
var
   nPosVirg: integer;
   cNumero: string;
begin
   nCasaDec := Min(10, nCasaDec);
   nCasaDec := Max(0, nCasaDec);

   Result := FormatFloat(aNumCasa[nCasaDec], nNumero);

{   cNumero := FloatToStr(nNumero);
   nPosVirg := Pos(',', cNumero);
   if nPosVirg = 0 then
      Result := cNumero + ',' + StringOfChar('0',nCasaDec)
   else
      Result := Copy(cNumero,1,nPosVirg+nCasaDec);}
end;

function Arredonda2(nNumero: Extended; nCasaDec: Integer): Extended;
// http://www.latiumsoftware.com/en/delphi/00033.php
// RoundD(123.456, 0) = 123.00
// RoundD(123.456, 2) = 123.46
// RoundD(123456, -3) = 123000
var
  n: Extended;
begin
  n := IntPower(10, nCasaDec);
  nNumero := nNumero * n;
  Result := (Int(nNumero) + Int(Frac(nNumero) * 2)) / n;
end;

function DiasUteis(DataIni:TDate; nDias: integer ): TDate;
var
   a : integer;
begin
   for a := 1 to nDias do
   begin
      DataIni := DataIni + 1;
      If DayOfTheWeek(DataIni) = 6 then
         DataIni := DataIni + 1;
      If DayOfTheWeek(DataIni) = 7 then
         DataIni := DataIni + 1;
   end;
   Result := DataIni;
end;

function IdGestC( DataIni:TDateTime; DataFim:TDateTime ): String;
var
   nTotal, nDias, nSeman : integer;
begin
   Result := '??sem?/7';
   if (DateToStr(DataIni) <> '  /  /    ') and ( DataIni < DataFim ) then
   begin
      nTotal := Round(DataFim - DataIni);
      nDias := nTotal mod 7;
      nSeman := Round((nTotal-nDias)/7);
      if nDias <> 0 then
         Result :=  IntToStr(nSeman)+'sem'+IntToStr(nDias)+'/7'
      else
         Result := IntToStr(nSeman)+'sem';
   end;
end;

function IdGestN( DataIni:TDateTime; DataFim:TDateTime ): single;
var
   nTotal, nDias, nSeman : integer;
begin
   Result := 0;
   if (DateToStr(DataIni) <> '  /  /    ') and ( DataIni < DataFim ) then
   begin
      nTotal := Round(DataFim - DataIni);
      nDias := nTotal mod 7;
      nSeman := Round((nTotal-nDias)/7);
      if nDias <> 0 then
         Result := nSeman + (nDias/7)
      else
         Result := nSeman;
   end;
end;

function IdGestNtoC( nIdGest: single ): String;
const
   nDecSem: single = 1.42857142;
var
   nSem, nDias: integer;
begin
   Result := '??sem?/7';
   nSem := Trunc(nIdGest);
   nDias := Round(((nIdGest-nSem)*10)/nDecSem);
   if nDias = 0 then
      Result := IntToStr(nSem)+'sem'
   else if nDias < 7 then
      Result := IntToStr(nSem)+'sem'+IntToStr(nDias)+'/7'
   else
      Result := IntToStr(nSem+1)+'sem';
end;

function IdGestCtoN( cIdGest: string ): integer;
var
   cDias, cSem : string;
   nPos, nSem, nDias: integer;
   dData : TDate;
begin
   Result := 0;

   cIdGest := BuscaTroca(cIdGest,' ','0');
   dData := 0;

   nPos := Pos('/7',cIdGest);
   if nPos <> 0 then
      cDias := Copy(cIdGest,nPos-1,1);

   if Trim(cDias) <> '' then
      nDias := StrToInt(cDias)
   else
      nDias := 0;

   nPos := Pos('sem', cIdGest);
   if nPos <> 0 then
      cSem := Before(cIdGest,'sem');

   if (cSem <> ' ') and (cSem <> '  ') then
       nDias := nDias + Round(StrToInt(cSem)*7);

   dData := dData + nDias;

   Result := Round(dData);
end;

function DataIGCDat(cIgest: String; DataIni: TDateTime): string;
var
   nPosDia, nPosSem, nDias: integer;
   cDias, cSem: string;
   dDum: TDateTime;
begin
   Result := '00sem0/7';
   cIGest := BuscaTroca(cIGest,' ','0');
   dDum := StrToDate('01/01/0001');

   nPosDia := Pos('/7',cIGest);
   if nPosDia <> 0 then
      cDias := Copy(cIGest,nPosDia-1,1);

   if cDias <> ' ' then
      nDias := StrToInt(cDias)
   else
      nDias := 0;

   nPosSem := Pos('sem', cIGest);
   if nPosSem <> 0 then
      cSem := Before(cIGest,'sem');

   if (cSem <> ' ') and (cSem <> '  ') then
       nDias := nDias + Round(StrToInt(cSem)*7);

//   if DataIni <> StrToDate('  /  /    ') then
      dDum := DataIni + Round(-1*nDias);

   if dDum <> StrToDate('01/01/0001') then
      Result := IdGestC(dDum, Now);
end;

function exp2(X: Single): Single;
{ Wichtig: im RC - Feld muss rcNearOrEven stehen, damit das
  Argument von f2xm1 sich im erlaubten Bereich befindet }
assembler;
const
   { FPU - Statusregister }
   stC0		= $0100;
var
   StatusWord: word;
asm
	fld	X
	frndint
	fld	X
	fsub	st, st(1)
	ftst
	fstsw	StatusWord
	fwait
	test	StatusWord, stC0
	jz	@ge
	fabs
@ge:	f2xm1	{ FPU >= 387: -1 <= st <= 1; FPU <= 287 : 0 <= st <= 0.5 }
	fld1
	faddp	st(1), st
	jz	@ge2
	fld1
	fdivr
@ge2:	fscale	{ FPU <= 287 : Resultat undefiniert, falls 0 < st(1) < 1 }
	fstp	st(1)
	fwait
end;

function BuscaArquivos(cDiretorio: string; cExt: string ): TStringList;
var
  SearchRec: TSearchRec;
  aArquivos : array of string;
  ListaArq : TStringList;
  nArqs : integer;
begin
   nArqs := 0;
   SetLength(aArquivos,nArqs);
   ListaArq := TStringList.Create;

   if RightStr(cDiretorio, 1) <> '\' then
      cDiretorio := cDiretorio+'\';

   // Verifica presenÁa de arquivos
   if FindFirst(cDiretorio+cExt, faAnyFile, SearchRec) = 0 then
   begin
      Inc(nArqs);
      SetLength(aArquivos,nArqs);
      aArquivos[nArqs-1] := SearchRec.Name;
      ListaArq.Add(SearchRec.Name);

      while (FindNext(SearchRec) = 0) do
      begin
         Inc(nArqs);
         SetLength(aArquivos,nArqs);
         aArquivos[nArqs-1] := SearchRec.Name;
      ListaArq.Add(SearchRec.Name)
      end;
      FindClose(SearchRec);
   end;
   Result := ListaArq;
end;

function BuscaVetorStr(StrVetor: array of string; SubStr: string): integer;
var
   a : integer;
begin
   Result := -1;
   for a := Low(StrVetor) to High(StrVetor) do
      if StrVetor[a] = SubStr then
         Result := a;

end;

function BuscaVetorInt(StrVetor: array of integer; SubStr: integer): integer;
var
   a : integer;
begin
   Result := -1;
   for a := Low(StrVetor) to High(StrVetor) do
      if StrVetor[a] = SubStr then
         Result := a;
end;

function BuscaVetorReal(var StrVetor: array of Real; SubStr: Real): Real;
var
   a : integer;
begin
   Result := -1;
   for a := Low(StrVetor) to High(StrVetor) do
      if StrVetor[a] = SubStr then
         Result := a;

end;

function BiosDate: String;
// Retorna a data da fabricaÁ„o do Chip da Bios do sistema
var
   Reg : TRegistry;
   cBiosDt : string;
begin
   cBiosDt := '??????????';
   try
      Reg := TRegistry.Create;
      Reg.RootKey := HKEY_LOCAL_MACHINE;
      if Reg.OpenKey('\HARDWARE\DESCRIPTION\System', False) then
         cBiosDt := Reg.ReadString('SystemBiosDate')
      else
         cBiosDt := string(pchar(ptr($FFFF5)));
   except
   end;
   Result := cBiosDt;
end;

function GetBuildInfo:string;
var
   VerInfoSize:  DWORD;
   VerInfo:      Pointer;
   VerValueSize: DWORD;
   VerValue:     PVSFixedFileInfo;
   Dummy:        DWORD;
   V1, V2, V3, V4: Word;
   Prog : string;
begin
   Prog := Application.Exename;
   VerInfoSize := GetFileVersionInfoSize(PChar(prog), Dummy);
   GetMem(VerInfo, VerInfoSize);
   GetFileVersionInfo(PChar(prog), 0, VerInfoSize, VerInfo);
   VerQueryValue(VerInfo, '\', Pointer(VerValue), VerValueSize);
   with VerValue^ do
        begin
        V1 := dwFileVersionMS shr 16;
        V2 := dwFileVersionMS and $FFFF;
        V3 := dwFileVersionLS shr 16;
        V4 := dwFileVersionLS and $FFFF;
        end;
   FreeMem(VerInfo, VerInfoSize);
   result := Copy (IntToStr (100 + v1), 3, 2) + '.' +
             Copy (IntToStr (100 + v2), 3, 2) + '.' +
             Copy (IntToStr (100 + v3), 3, 2) + '.' +
             Copy (IntToStr (100 + v4), 3, 2);
end;

function WinExecAndWait32( FileName: String; Visibility  : Integer ) : Cardinal;
// Resulta -1 se o processo falhar
var { by Pat Ritchey }
   zAppName     : array[0..512] of char;
   zCurDir      : array[0..255] of char;
   WorkDir      : String;
   StartupInfo  : TStartupInfo;
   ProcessInfo  : TProcessInformation;
begin
   StrPCopy( zAppName, FileName );
   GetDir  ( 0,        WorkDir  );
   StrPCopy( zCurDir,  WorkDir  );

   FillChar( StartupInfo, Sizeof( StartupInfo ), #0 );

   StartupInfo.cb          := Sizeof( StartupInfo );
   StartupInfo.dwFlags     := STARTF_USESHOWWINDOW;
   StartupInfo.wShowWindow := Visibility;

   if ( not CreateProcess( nil,
                           zAppName, { pointer to command line string }
                           nil, { pointer to process security attributes }
                           nil, { pointer to thread security attributes }
                           false, { handle inheritance flag }
                           CREATE_NEW_CONSOLE or { creation flags }
                           NORMAL_PRIORITY_CLASS,
                           nil, { pointer to new environment block }
                           zCurDir, { pointer to current directory name }
                           StartupInfo, { pointer to STARTUPINFO }
                           ProcessInfo ) ) then
   begin
       Result := $FFFFFFFF; { pointer to PROCESS_INF }
       MessageBox( Application.Handle, PChar( SysErrorMessage( GetLastError ) ), 'Yipes!', 0 );
   end
   else
   begin
       WaitforSingleObject( ProcessInfo.hProcess, INFINITE );
       GetExitCodeProcess ( ProcessInfo.hProcess, Result   );
       CloseHandle        ( ProcessInfo.hProcess           );
       CloseHandle        ( ProcessInfo.hThread            );
   end;
end;

procedure PostKeyEx32(Key: Word; const Shift: TShiftState; SpecialKey: boolean);
// Exemplo : PostKeyEx32(Ord('A'), [ssCtrl], false);
// Envia Ctrl+A para o controle que tiver o foco.
// Key : virtual keycode da tecla a enviar. Para caracteres
// imprimÌveis informe o cÛdigo ANSI (Ord(CHARACTER)).
// Shift : estado das teclas modificadoras.
// Shift, Control, Alt, Mouse Buttons.
// SpecialKey: normalmente deve ser False. Informe True se
// a tecla desejada for, por exemplo, do teclado numÈrico.
type
   TShiftKeyInfo = Record
                     shift: Byte;
                     vkey : Byte;
                   End;
   byteset = Set of 0..7;
const  ShiftKeys: array [1..3] of TShiftKeyInfo =
   ((shift: Ord(ssCtrl); vkey: VK_CONTROL ),
   (shift: Ord(ssShift); vkey: VK_SHIFT ),
   (shift: Ord(ssAlt); vkey: VK_MENU ));
var
   Flag: DWORD;
   bShift: ByteSet absolute shift;
   i: Integer;
begin
   for i := 1 to 3 do
   begin
      if shiftkeys[i].shift in bShift then
         Keybd_Event(ShiftKeys[i].vkey,
            MapVirtualKey(ShiftKeys[i].vkey, 0), 0, 0);
   end; // for

   if SpecialKey Then
      Flag := KEYEVENTF_EXTENDEDKEY
   else
      Flag := 0;

   Keybd_Event(Key, MapvirtualKey(Key, 0), Flag, 0);
   Flag := Flag or KEYEVENTF_KEYUP;
   Keybd_Event(Key, MapvirtualKey(Key, 0), Flag, 0);

   for i := 3 DownTo 1 do
   begin
      if ShiftKeys[i].shift in bShift then
         Keybd_Event(shiftkeys[i].vkey,
            MapVirtualKey(ShiftKeys[i].vkey, 0),
            KEYEVENTF_KEYUP, 0);
   end; // for
end; // PostKeyEx32

function AlteraDataHora (dDataHora : TDateTime) : boolean ;
var
   st: TSystemTime;
   Ano, Mes, Dia, Hora, Minuto, Segundo, MSeg: word;
begin
   DecodeDate(dDataHora, Ano, Mes, Dia);
   DecodeTime(dDataHora, Hora, Minuto, Segundo, MSeg);

   GetLocalTime(st);
   st.wYear := Ano;
   st.wMonth := Mes;
   st.wDay := Dia;
   st.wHour := Hora;
   st.wMinute := Minuto;
   st.wSecond := Segundo;
   if not SetLocalTime(st) then
      Result := False
   else
      Result := True;
end;

function DriveOk(Drive: string): boolean;
const
   cDrive : string = 'ABCDEFGHIJKLMNOPQRSTUVXWYZ';
var
   I: byte;
begin
   Result := False;
   Drive := UpperCase(Drive);
   I := Pos(Drive, cDrive);
   if I = 0 then
      raise Exception.Create('Unidade incorreta');

   Result := (DiskSize(I) >= 0);
end;

function DriveSpace(Drive: string): integer;
const
   cDrive : string = 'ABCDEFGHIJKLMNOPQRSTUVXWYZ';
var
   I: byte;
begin
   Result := 0;
   Drive := UpperCase(Drive);
   I := Pos(Drive, cDrive);
   if I = 0 then
      raise Exception.Create('Unidade incorreta');

   Result := DiskFree(I);
end;

procedure AddHexString(S : String; Lines : TStrings );
var AddS, HexS, CopyS : String;
    i : Integer;
const SLen = 8;
begin
  while Length(S) > 0 do
    begin
      AddS := Copy(S,1,SLen);
      HexS := '';
      Delete(S,1,SLen);
      for i := 1 to SLen do
        begin
          CopyS := Copy(AddS,i,1);
          if CopyS <> '' then
            HexS := HexS + ' ' + Format('%2.2x',[Byte(CopyS[1])]) //
          else
            HexS := HexS + '   ';
        end;
       while Length(AddS) < SLen do
         AddS := AddS + ' ';
       for i := 1 to SLen do
         case AddS[i] of
           #0..#31 : AddS[i] := '.';
           #127    : AddS[i] := '.';
         end;
       Lines.Add(HexS+' : '+AddS);
    end;
end;

function CaptureScreenRect( ARect: TRect ): TJPEGImage;
//Image1.picture.Assign(CaptureScreenRect(Rect(0,0,Width,Height)));
var
   ScreenDC: HDC;
   ImageBMP : TBitmap;
begin
   ImageBMP := TBitmap.Create;
   Result := TJPEGImage.Create;
   with ImageBMP, ARect do
   begin
      Width := Right - Left;
      Height := Bottom - Top;
      ScreenDC := GetDC( 0 );
      try
         BitBlt( Canvas.Handle, 0, 0, Width, Height, ScreenDC, Left, Top, SRCCOPY );
      finally
         ReleaseDC( 0, ScreenDC );
      end;
      // Palette := GetSystemPalette;
   end;
   with Result do
   begin
      CompressionQuality := 20;
      Grayscale := False; // False: melhor e maior, True: pior e menor
      PixelFormat := jf24Bit; // ou jf8Bit
      Assign(ImageBMP);
      Compress;
//      SaveToFile(ChangeFileExt(NomeArq,'.jpg'));
      ImageBMP.Free;
   end;
end;

function CompactaBD(ArqOrigem, SenOrigem, ArqDestino, SenDestino : string ; lInfo : boolean) : boolean;
const
   SProvider = 'Provider=Microsoft.Jet.OLEDB.4.0;'+
               'Jet OLEDB:Engine Type=5;'+   // Jet 4.0
               'Data Source=';
var
   nBD : integer;

   JE : TJetEngine; //Jet Engine
   cMsg : String;
begin
   try
      Result := False;
      cMsg := 'Falha na compactaÁ„o !';
      if FileExists(ArqDestino) then
         DeleteFile(ArqDestino);

      if FileExists(ArqOrigem) and
         not FileExists(BuscaTroca(ArqOrigem,'.mdb','.ldb')) then
      begin
         JE:= TJetEngine.Create(Application);
         try
            JE.CompactDatabase(SProvider+ArqOrigem+
               ';Mode=Share Exclusive;'+
               'Jet OLEDB:Database Password='+SenOrigem,
                               SProvider+ArqDestino+
               ';Mode=Share Exclusive;'+
               'Jet OLEDB:Database Password='+SenDestino);

            lInfo := False;
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
      if lInfo then
         ShowMessage(cMsg);
   end;
end;

procedure SetDefaultPrinter(PrinterName: String);
var
  nImp : Integer;
  Device : PChar;
  Driver : Pchar;
  Port : Pchar;
  HdeviceMode: Thandle;
  aPrinter : TPrinter;
begin
//  Printer.PrinterIndex := -1;
  getmem(Device, 255);
  getmem(Driver, 255);
  getmem(Port, 255);
  aPrinter := TPrinter.Create;
  nImp := Printer.Printers.IndexOf(PrinterName);
  if nImp > -1 then
  begin
     aPrinter.PrinterIndex := nImp;
     aPrinter.GetPrinter(device, driver, port, HdeviceMode);
     StrCat(Device, ',');
     StrCat(Device, Driver );
     StrCat(Device, Port );
     WriteProfileString('windows', 'device', Device);
     StrCopy( Device, 'windows' );
     SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, Longint(@Device));
  end;

  Freemem(Device, 255);
  Freemem(Driver, 255);
  Freemem(Port, 255);

  FreeAndNil(aPrinter);

  Sleep(1000);
end;

function ReparaMDB(cNomeMDB : String; lAutom : boolean) : Boolean;
var
   dao: Variant;
   nBD : integer;
begin
   Result := False;

   if FileExists(cNomeMDB) then
   try
//      Screen.Cursor := crHourGlass;

      { Get a handle  to the Database  Engine. }
      dao := CreateOleObject('DAO.DBEngine.36');
      { Repair the  database.  }
      dao.RepairDatabase(cNomeMDB);

      Result := True;

//      Screen.Cursor  := crDefault;
      if not lAutom then
         Application.MessageBox (PChar('Data base repaired !'#13#10+
               cNomeMDB),
               'ConcluÌdo',MB_OK+MB_ICONEXCLAMATION);
   except
      on E: EOleException do
      begin
//         Screen.Cursor  := crDefault;
         if not lAutom then
            Application.MessageBox (PChar('The system encontered an error to repair data base:'#13#10+
                  cNomeMDB+'. Try again !'+#13#10+
                  E.Message),
                  'Erro',MB_OK+MB_ICONERROR);
      end;
   end;  // try
end;

function Maiuscula(Texto:String): String;
begin
   Result := BuscaTroca(Texto,'·','¡');
   Result := BuscaTroca(Result,'‚','¬');
   Result := BuscaTroca(Result,'„','√');
   Result := BuscaTroca(Result,'È','…');
   Result := BuscaTroca(Result,'Í',' ');
   Result := BuscaTroca(Result,'Ì','Õ');
   Result := BuscaTroca(Result,'Û','”');
   Result := BuscaTroca(Result,'Ù','‘');
   Result := BuscaTroca(Result,'ı','’');
   Result := BuscaTroca(Result,'˙','⁄');
   Result := BuscaTroca(Result,'¸','‹');
   Result := BuscaTroca(Result,'Á','«');
   Result := UpperCase(Result);
end;

function PrimaPalavra(Texto:String; lLimpa: boolean): String;
const
  accent   : string = '„‡·‰‚ËÈÎÍÏÌÔÓıÚÛˆÙ˘˙¸˚Á√¿¡ƒ¬»…À ÃÕœŒ’“”÷‘Ÿ⁄‹€«';
  noaccent  : string = 'aaaaaeeeeiiiiooooouuuucAAAAAEEEEIIIIOOOOOUUUUC';
var
   i : integer;
begin
   Texto := Copy(Texto, 1, Pos(' ', Texto)-1);

   if lLimpa then
      for i := 1 to length(accent) do
         while Pos(accent[i],Texto) > 0 do
            Texto[Pos(accent[i],Texto)] := noaccent[i];

   Result := Texto;
end;

function Mes(dData : TDate) : integer;
var
   nAno,nMes,nDia: Word;
begin
   DecodeDate(dData, nAno, nMes, nDia);
   Result := nMes;
end;

function Ano(dData : TDate) : integer;
var
   nAno,nMes,nDia: Word;
begin
   DecodeDate(dData, nAno, nMes, nDia);
   Result := nAno;
end;

function DateTimeToSQLDateTimeString(Data: TDateTime; Format: string;
  OnlyDate: Boolean = True): string;
var
  y, m, d, h, mm, s, ms: Word;
begin
  DecodeDate(Data, y, m, d);
  DecodeTime(Data, h, mm, s, ms);
  if Format = 'dmy' then
    Result := IntToStr(d) + '-' + IntToStr(m) + '-' + IntToStr(y)
  else if Format = 'ymd' then
    Result := IntToStr(y) + '-' + IntToStr(m) + '-' + IntToStr(d)
  else if Format = 'ydm' then
    Result := IntToStr(y) + '-' + IntToStr(d) + '-' + IntToStr(m)
  else if Format = 'myd' then
    Result := IntToStr(m) + '-' + IntToStr(y) + '-' + IntToStr(d)
  else if Format = 'dym' then
    Result := IntToStr(d) + '-' + IntToStr(y) + '-' + IntToStr(m)
  else
    Result := IntToStr(m) + '-' + IntToStr(d) + '-' + IntToStr(y); //mdy: ; //US
  if not OnlyDate then
    Result := Result + ' ' + IntToStr(h) + ':' + IntToStr(mm) + ':' + IntToStr(s);
end;

function ValidaCPF(Const S: ShortString): Boolean;
var
 i, Numero, Resto: Integer;
 DV1, DV2: Byte;
 Total, Soma: Integer;
 cCPF : string;
begin
   result := FALSE;
   cCPF := BuscaTroca(s, '.', '');
   cCPF := BuscaTroca(cCPF, '-', '');
   cCPF := BuscaTroca(cCPF, '/', '');
   if s = '11111111111' then
   else if length(cCPF) = 11 then
   begin
      Total := 0 ;
      Soma := 0 ;
      for i := 1 to 9 do
      begin
      try
         Numero := StrToInt(cCPF[I]);
      except
         Numero := 0;
      end;
         Total := Total + (Numero * ( 11 - i));
         Soma := Soma + Numero;
      end;
      if Total > 0 then
      begin
         Resto := Total mod 11 ;
         if Resto > 1 then
            DV1 := 11 - Resto
         else
            DV1 := 0 ;
         Total := Total + Soma + 2 * DV1;
         Resto := Total mod 11 ;
         if Resto > 1 then
            DV2 := 11 - Resto
         else
            DV2 := 0 ;
         if (DV1 = StrToInt(cCPF[ 10 ])) and
               (DV2 = StrToInt(cCPF[ 11 ])) then
            result := TRUE;
      end;
   end;
end;

function ValidaCNPJ(const s: string): boolean;
var
   i, soma, mult: integer;
   cCNPJ : string;
begin
   Result := False;
   cCNPJ := BuscaTroca(s, '.', '');
   cCNPJ := BuscaTroca(cCNPJ, '-', '');
   cCNPJ := BuscaTroca(cCNPJ, '/', '');
   if Length(cCNPJ) <> 14 then
      Exit;
   soma := 0;
   mult := 2;
   for i := 12 downto 1 do
   begin
      soma := soma + StrToInt(cCNPJ[i]) * mult;
      mult := mult + 1;
      if mult > 9 then
         mult := 2;
   end;
   mult := soma mod 11;
   if mult <= 1 then
      mult := 0
   else
      mult := 11 - mult;
   if mult <> StrToInt(cCNPJ[13]) then
      Exit;
   soma := 0;
   mult := 2;
   for i := 13 downto 1 do
   begin
      soma := soma + StrToInt(cCNPJ[i]) * mult;
      mult := mult + 1;
      if mult > 9 then
         mult := 2;
   end;
   mult := soma mod 11;
   if mult <= 1 then
      mult := 0
   else
      mult := 11 - mult;
   Result := (mult = StrToInt(cCNPJ[14]));
end;

function ValidaEMail2(const EMailIn: string):Boolean;
const
  CaraEsp: array[1..40] of string[1] =
  ( '!','#','$','%','®','&','*',
  '(',')','+','=','ß','¨','¢','π','≤',
  '≥','£','¥','`','Á','«',',',';',':',
  '<','>','~','^','?','/','','|','[',']','{','}',
  '∫','™','∞');
var
  i,cont   : integer;
  EMail    : ShortString;
begin
  EMail := EMailIn;
  Result := True;
  cont := 0;
  if EMail <> '' then
    if (Pos('@', EMail)<>0) and (Pos('.', EMail)<>0) then    // existe @ .
    begin
      if (Pos('@', EMail)=1) or (Pos('@', EMail)= Length(EMail)) or (Pos('.', EMail)=1) or (Pos('.', EMail)= Length(EMail)) or (Pos(' ', EMail)<>0) then
        Result := False
      else                                   // @ seguido de . e vice-versa
        if (abs(Pos('@', EMail) - Pos('.', EMail)) = 1) then
          Result := False
        else
          begin
            for i := 1 to 40 do            // se existe Caracter Especial
              if Pos(CaraEsp[i], EMail)<>0 then
                Result := False;
            for i := 1 to length(EMail) do
            begin                                 // se existe apenas 1 @
              if EMail[i] = '@' then
                cont := cont + 1;                    // . seguidos de .
              if (EMail[i] = '.') and (EMail[i+1] = '.') then
                Result := false;
            end;
                                   // . no f, 2ou+ @, . no i, - no i, _ no i
            if (cont >=2) or ( EMail[length(EMail)]= '.' )
              or ( EMail[1]= '.' ) or ( EMail[1]= '_' )
              or ( EMail[1]= '-' )  then
                Result := false;
                                            // @ seguido de COM e vice-versa
            if (abs(Pos('@', EMail) - Pos('com', EMail)) = 1) then
              Result := False;
                                              // @ seguido de - e vice-versa
            if (abs(Pos('@', EMail) - Pos('-', EMail)) = 1) then
              Result := False;
                                              // @ seguido de _ e vice-versa
            if (abs(Pos('@', EMail) - Pos('_', EMail)) = 1) then
              Result := False;
          end;
    end
    else
      Result := False;
end;

function ValidaEmail(email: string): boolean;
  // Returns True if the email address is valid
  // Author: Ernesto D'Spirito
const
    // Valid characters in an "atom"
    atom_chars = [#33..#255] - ['(', ')', '<', '>', '@', ',', ';', ':',
                                '\', '/', '"', '.', '[', ']', #127];
    // Valid characters in a "quoted-string"
    quoted_string_chars = [#0..#255] - ['"', #13, '\'];
    // Valid characters in a subdomain
    letters = ['A'..'Z', 'a'..'z'];
    letters_digits = ['0'..'9', 'A'..'Z', 'a'..'z'];
    subdomain_chars = ['-', '0'..'9', 'A'..'Z', 'a'..'z'];
type
    States = (STATE_BEGIN, STATE_ATOM, STATE_QTEXT, STATE_QCHAR,
      STATE_QUOTE, STATE_LOCAL_PERIOD, STATE_EXPECTING_SUBDOMAIN,
      STATE_SUBDOMAIN, STATE_HYPHEN);
var
    State: States;
    i, n, subdomains: integer;
    c: char;
begin
    State := STATE_BEGIN;
    n := Length(email);
    i := 1;
    subdomains := 1;
    while (i <= n) do begin
      c := email[i];
      case State of
      STATE_BEGIN:
        if c in atom_chars then
          State := STATE_ATOM
        else if c = '"' then
          State := STATE_QTEXT
        else
          break;
      STATE_ATOM:
        if c = '@' then
          State := STATE_EXPECTING_SUBDOMAIN
        else if c = '.' then
          State := STATE_LOCAL_PERIOD
        else if not (c in atom_chars) then
          break;
      STATE_QTEXT:
        if c = '\' then
          State := STATE_QCHAR
        else if c = '"' then
          State := STATE_QUOTE
        else if not (c in quoted_string_chars) then
          break;
      STATE_QCHAR:
        State := STATE_QTEXT;
      STATE_QUOTE:
        if c = '@' then
          State := STATE_EXPECTING_SUBDOMAIN
        else if c = '.' then
          State := STATE_LOCAL_PERIOD
        else
          break;
      STATE_LOCAL_PERIOD:
        if c in atom_chars then
          State := STATE_ATOM
        else if c = '"' then
          State := STATE_QTEXT
        else
          break;
      STATE_EXPECTING_SUBDOMAIN:
        if c in letters then
          State := STATE_SUBDOMAIN
        else
          break;
      STATE_SUBDOMAIN:
        if c = '.' then begin
          inc(subdomains);
          State := STATE_EXPECTING_SUBDOMAIN
        end else if c = '-' then
          State := STATE_HYPHEN
        else if not (c in letters_digits) then
          break;
      STATE_HYPHEN:
        if c in letters_digits then
          State := STATE_SUBDOMAIN
        else if c <> '-' then
         break;
      end;
      inc(i);
    end;
    if i <= n then
      Result := False
    else
      Result := (State = STATE_SUBDOMAIN) and (subdomains >= 2);
end;

{
type
   TForm1 = class(TForm)
     Button1: TButton;
     procedure Button1Click(Sender: TObject) ;
     procedure CallMeByName(Sender: TObject) ;
   private
     procedure ExecMethod(OnObject: TObject; MethodName: string) ;
   end;

var
   Form1: TForm1;

type
   TExec = procedure of object;

procedure TForm1.CallMeByName(Sender: TObject) ;
begin
   ShowMessage('Hello Delphi!') ;
end;

procedure TForm1.Button1Click(Sender: TObject) ;
begin
   ExecMethod(Form1, 'CallMeByName') ;
end; }

procedure ExecMethod(OnObject: TObject; MethodName: string) ;
var
   Routine: TMethod;
   Exec: TExec;
begin
   Routine.Data := Pointer(OnObject) ;
   Routine.Code := OnObject.MethodAddress(MethodName) ;
   if NOT Assigned(Routine.Code) then Exit;
   Exec := TExec(Routine) ;
   Exec;
end;

function QuotedData(dData : TDateTime; nTpBD : integer) : string;
begin
   if nTpBD = 1 then
      Result := '#'+FormatDateTime('mm/dd/yyyy', dData)+'#'
   else if nTpBD = 2 then
      Result := QuotedStr(DateTimeToStr(dData));
end;

function QuotedDataHora(dData : TDateTime; nTpBD : integer;  cFormato : string = 'dd/mm/yyyy hh:nn:ss') : string;
begin
   if nTpBD = 1 then
   begin
      // inverte data para o access
      cFormato := BuscaTroca(cFormato, 'dd/mm', 'mm/dd');
      Result := '#'+FormatDateTime(cFormato, dData)+'#';
   end
   else if nTpBD = 2 then
      Result := QuotedStr(FormatDateTime(cFormato, dData));
end;

function QuotedTexto(cTexto : string; nTpBD : integer) : string;
begin
//   cTexto := BuscaTroca(cTexto, ':', ';');
   if nTpBD = 1 then
   begin
      cTexto := BuscaTroca(cTexto, '"', '`');
      Result := '"'+cTexto+'"';
   end
   else if nTpBD = 2 then
      Result := QuotedStr(cTexto);
end;

function QuotedBool(lBoll : boolean; nTpBD : integer) : string;
begin
   if nTpBD = 1 then
      Result := BoolToStr(lBoll, True)
   else if nTpBD = 2 then
      Result := QuotedStr(BoolToStr(lBoll, True));
end;

function QuotedMoney(cValor : string; nTpBD : integer) : string;
begin
   Result := BuscaTroca(cValor, '.', '');
   Result := BuscaTroca(Result, ',', '.');
end;

function NomeMaquina: string;
var
   Reg: TRegistry;
   buffer : array[0..255] of char;
   size : dword;
   UserName: String;
   ComputerName: String;
begin
   try
      size := MAX_COMPUTERNAME_LENGTH + 1;
      GetComputerName(buffer, size);
      ComputerName := buffer;
      Result := ComputerName;
   except
      Result := 'Maquina0';
   end;

   if Result = 'Maquina0' then
   try
      Reg := TRegistry.Create;
      Reg.RootKey := HKEY_LOCAL_MACHINE;
      if Reg.KeyExists('\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName') then
      begin
         // Para Windows 98 / NT / XP
         Reg.OpenKey('\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName', False);
         Result := Reg.ReadString('ComputerName');
      end
      else if Reg.KeyExists('\HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\MS SETUP (ACME)') then
      begin
         // Para Windows 95
         Reg.RootKey := HKEY_CURRENT_USER;
         Reg.OpenKey('SOFTWARE\MICROSOFT\MS SETUP (ACME)\USER NAME', False);
         Result := Reg.ReadString('DefName');
      end;
   finally
      Reg.free;
   end;
end;

function IdWin : string;
var
   Reg: TRegistry;
{   binarySize: INTEGER;
   HexBuf: array of BYTE;
   function DecodeProductKey(const HexSrc: array of Byte): string;
   const
     StartOffset: Integer = $34;  //Offset 34 = Array[52]
     EndOffset: Integer   = $34 + 15; //Offset 34 + 15(Bytes) = Array[64]
     Digits: array[0..23] of CHAR = ('B', 'C', 'D', 'F', 'G', 'H', 'J',
       'K', 'M', 'P', 'Q', 'R', 'T', 'V', 'W', 'X', 'Y', '2', '3', '4', '6', '7', '8', '9');
     dLen: Integer = 29;  //Length of Decoded Product Key
     sLen: Integer = 15;
     //Length of Encoded Product Key in Bytes (An total of 30 in chars)
   var
     HexDigitalPID: array of CARDINAL;
     Des: array of CHAR;
     I, N: INTEGER;
     HN, Value: CARDINAL;
   begin
     SetLength(HexDigitalPID, dLen);
     for I := StartOffset to EndOffset do
     begin
       HexDigitalPID[I - StartOffSet] := HexSrc[I];
     end;

     SetLength(Des, dLen + 1);

     for I := dLen - 1 downto 0 do
     begin
       if (((I + 1) mod 6) = 0) then
       begin
         Des[I] := '-';
       end
       else
       begin
         HN := 0;
         for N := sLen - 1 downto 0 do
         begin
           Value := (HN shl 8) or HexDigitalPID[N];
           HexDigitalPID[N] := Value div 24;
           HN    := Value mod 24;
         end;
         Des[I] := Digits[HN];
       end;
     end;
     Des[dLen] := Chr(0);

     for I := 0 to Length(Des) do
     begin
       Result := Result + Des[I];
     end;
   end; }
begin
   Result := '';
   Reg := TRegistry.Create;
   Reg.RootKey := HKEY_LOCAL_MACHINE;
   Reg.OpenKey('\Software\MicroSoft\Windows NT\CurrentVersion', False);
{  // pega chave de instalaÁ„o do windows
   if Reg.GetDataType('DigitalProductId') = rdBinary then
   try
     binarySize := Reg.GetDataSize('DigitalProductId');
     SetLength(HexBuf, binarySize);
     if binarySize > 0 then
     begin
       Reg.ReadBinaryData('DigitalProductId', HexBuf[0], binarySize);
       Result := DecodeProductKey(HexBuf);
     end;   // 'TY74Y-2C83J-4W373-XYKV2-689MB'#0#0
   except
   end;
}   if Result = '' then   // '76501-OEM-0011903-00128'
      Result := Reg.ReadString('ProductId');
   if Result = '' then
   begin
      Reg.OpenKey('\Software\MicroSoft\Windows\CurrentVersion', False);
      Result := Reg.ReadString('ProductId');
      if Result = '' then
         Result := 'ProductId n„o Indentificado';
   end;
   Reg.Free;
end;

function GetCommonPathFiles : String;
// Pega o diretorio de Arquivos Comuns
var
   Reg : TRegistry;
begin
   Reg := TRegistry.Create;
   Reg.RootKey := HKEY_LOCAL_MACHINE;
   if reg.OpenKey('\software\microsoft\windows\currentversion\',False) Then
      Result := Reg.ReadString('CommonFilesDir');
   if not Reg.ValueExists('CommonFilesDir') then
      Result := '';
   if Copy(Result,Length(Result),1)<>'/' then
      Result := Result + '\';
   Reg.CloseKey;
   Reg.Free;
end;

function ListarArquivos(Diretorio: string; Sub:Boolean) : TStrings;
var
  F: TSearchRec;
  Ret: Integer;
  TempNome: string;
  ListaDir : TStrings;
begin
   ListaDir := TStrings.Create;
   Result := ListaDir;
   Ret := FindFirst(Diretorio+'\*.*', faAnyFile, F);
   try
      while Ret = 0 do
      begin
         if TemAtributo(F.Attr, faDirectory) then
         begin
            if (F.Name <> '.') And (F.Name <> '..') then
               if Sub = True then
               begin
                  TempNome := Diretorio+'\' + F.Name;
                  ListaDir.AddStrings(ListarArquivos(TempNome, True));
               end;
         end
         else
            ListaDir.Add(Diretorio+'\'+F.Name);

         Ret := FindNext(F);
      end;
   finally
      Result := ListaDir;
      FreeAndNil(ListaDir);
      FindClose(F);
   end;
end;

function TemAtributo(Attr, Val: Integer): Boolean;
begin
   Result := Attr and Val = Val;
end;

Function DelTree(DirName : string): Boolean;
var
  SHFileOpStruct : TSHFileOpStruct;
  DirBuf : array [0..255] of char;
begin
  try
   Fillchar(SHFileOpStruct,Sizeof(SHFileOpStruct),0) ;
   FillChar(DirBuf, Sizeof(DirBuf), 0 ) ;
   StrPCopy(DirBuf, DirName) ;
   with SHFileOpStruct do begin
    Wnd := 0;
    pFrom := @DirBuf;
    wFunc := FO_DELETE;
    fFlags := FOF_ALLOWUNDO;
    fFlags := fFlags or FOF_NOCONFIRMATION;
    fFlags := fFlags or FOF_SILENT;
   end;
    Result := (SHFileOperation(SHFileOpStruct) = 0) ;
   except
    Result := False;
  end;
end;

procedure DelTree2(const Directory: TFileName);
{http://www.latiumsoftware.com/en/delphi/00016.php

 This code will remove the directory C:\TEMP\A123:
      DelTree('C:\TEMP\A123') }
var
  DrivesPathsBuff: array[0..1024] of char;
  DrivesPaths: string;
  len: longword;
  ShortPath: array[0..MAX_PATH] of char;
  dir: TFileName;
   procedure rDelTree(const Directory: TFileName);
   // Recursively deletes all files and directories
   // inside the directory passed as parameter.
   var
     SearchRec: TSearchRec;
     Attributes: LongWord;
     ShortName, FullName: TFileName;
     pname: pchar;
   begin
     if FindFirst(Directory + '*', faAnyFile and not faVolumeID,
        SearchRec) = 0 then begin
       try
         repeat // Processes all files and directories
           if SearchRec.FindData.cAlternateFileName[0] = #0 then
             ShortName := SearchRec.Name
           else
             ShortName := SearchRec.FindData.cAlternateFileName;
           FullName := Directory + ShortName;
           if (SearchRec.Attr and faDirectory) <> 0 then begin
             // It's a directory
             if (ShortName <> '.') and (ShortName <> '..') then
               rDelTree(FullName + '\');
           end else begin
             // It's a file
             pname := PChar(FullName);
             Attributes := GetFileAttributes(pname);
             if Attributes = $FFFFFFFF then
               raise EInOutError.Create(SysErrorMessage(GetLastError));
             if (Attributes and FILE_ATTRIBUTE_READONLY) <> 0 then
               SetFileAttributes(pname, Attributes and not
                 FILE_ATTRIBUTE_READONLY);
             if Windows.DeleteFile(pname) = False then
               raise EInOutError.Create(SysErrorMessage(GetLastError));
           end;
         until FindNext(SearchRec) <> 0;
       except
         FindClose(SearchRec);
         raise;
       end;
       FindClose(SearchRec);
     end;
     if Pos(#0 + Directory + #0, DrivesPaths) = 0 then begin
       // if not a root directory, remove it
       pname := PChar(Directory);
       Attributes := GetFileAttributes(pname);
       if Attributes = $FFFFFFFF then
         raise EInOutError.Create(SysErrorMessage(GetLastError));
       if (Attributes and FILE_ATTRIBUTE_READONLY) <> 0 then
         SetFileAttributes(pname, Attributes and not
           FILE_ATTRIBUTE_READONLY);
       if Windows.RemoveDirectory(pname) = False then begin
         raise EInOutError.Create(SysErrorMessage(GetLastError));
       end;
     end;
   end;
// ----------------
begin
  DrivesPathsBuff[0] := #0;
  len := GetLogicalDriveStrings(1022, @DrivesPathsBuff[1]);
  if len = 0 then
    raise EInOutError.Create(SysErrorMessage(GetLastError));
  SetString(DrivesPaths, DrivesPathsBuff, len + 1);
  DrivesPaths := Uppercase(DrivesPaths);
  len := GetShortPathName(PChar(Directory), ShortPath, MAX_PATH);
  if len = 0 then
    raise EInOutError.Create(SysErrorMessage(GetLastError));
  SetString(dir, ShortPath, len);
  dir := Uppercase(dir);
  rDelTree(IncludeTrailingBackslash(dir));
end;

function FindFile(const filespec: TFileName; attributes: integer): TStringList;
{
attributes = faReadOnly Or faHidden Or faSysFile Or faArchive

StringList := FindFile('C:\Delphi\*.pas')
}
var
  spec: string;
  list: TStringList;

   procedure RFindFile(const folder: TFileName);
   var
     SearchRec: TSearchRec;
   begin
     // Locate all matching files in the current
     // folder and add their names to the list
     if FindFirst(folder + spec, attributes, SearchRec) = 0 then begin
       try
         repeat
           if (SearchRec.Attr and faDirectory = 0) or
              (SearchRec.Name <> '.') and (SearchRec.Name <> '..') then
             list.Add(folder + SearchRec.Name);
         until FindNext(SearchRec) <> 0;
       except
         FindClose(SearchRec);
         raise;
       end;
       FindClose(SearchRec);
     end;
     // Now search the subfolders
     if FindFirst(folder + '*', attributes
         Or faDirectory, SearchRec) = 0 then
     begin
       try
         repeat
           if ((SearchRec.Attr and faDirectory) <> 0) and
              (SearchRec.Name <> '.') and (SearchRec.Name <> '..') then
             RFindFile(folder + SearchRec.Name + '\');
         until FindNext(SearchRec) <> 0;
       except
         FindClose(SearchRec);
         raise;
       end;
       FindClose(SearchRec);
     end;
   end; // procedure RFindFile inside of FindFile
begin // function FindFile
  list := TStringList.Create;
  try
    spec := ExtractFileName(filespec);
    RFindFile(ExtractFilePath(filespec));
    Result := list;
  except
    list.Free;
    raise;
  end;
end;

function ContaStr(Text, Busca : string) : integer;
{ Substitui um caractere dentro da string}
var
   nBusca, nPos, nConta : integer;
begin
   nConta := 0;
   nBusca := Length(Busca);
   nPos := Pos(Busca, Text);

   while (nPos <> 0) do
   begin
      Inc(nConta);
      Delete(Text,nPos,nBusca);
      nPos := Pos(Busca, Text);
   end;
   Result := nConta;
end;

procedure CriarAtalhoDT(FileName, Parameters, InitialDir,
                                 ShortcutName, ShortcutFolder : String);
//CRIAR ATALHO NO DESKTOP
var
   MyObject : IUnknown;
   MySLink : IShellLink;
   MyPFile : IPersistFile;
   Directory : String;
   WFileName : WideString;
   MyReg : TRegIniFile;
begin
   MyObject := CreateComObject(CLSID_ShellLink);
   MySLink := MyObject as IShellLink;
   MyPFile := MyObject as IPersistFile;
   with MySLink do
   begin
      SetArguments(PChar(Parameters));
      SetPath(PChar(FileName));
      SetWorkingDirectory(PChar(InitialDir));
   end;
   MyReg := TRegIniFile.Create('Software\MicroSoft\Windows\CurrentVersion\Explorer');
   Directory := MyReg.ReadString ('Shell Folders','Desktop','');
   WFileName := Directory + '\' + ShortcutName + '.lnk';
   MyPFile.Save (PWChar (WFileName), False);
   MyReg.Free;
end;

procedure CriarAtalhoIni(FileName, Parameters, InitialDir,
                                    ShortcutName, ShortcutFolder : String);
// CRIAR ATALHO NO MENU INICIAR
var
   MyObject : IUnknown;
   MySLink : IShellLink;
   MyPFile : IPersistFile;
   Directory : String;
   WFileName : WideString;
   MyReg : TRegIniFile;
begin
   MyObject := CreateComObject(CLSID_ShellLink);
   MySLink := MyObject as IShellLink;
   MyPFile := MyObject as IPersistFile;
   with MySLink do
   begin
      SetArguments(PChar(Parameters));
      SetPath(PChar(FileName));
      SetWorkingDirectory(PChar(InitialDir));
   end;
   MyReg := TRegIniFile.Create ('Software\MicroSoft\Windows\CurrentVersion\Explorer');
   Directory := MyReg.ReadString('Shell Folders','Start Menu','') + '\' + ShortcutFolder;
   CreateDir(Directory);
   WFileName := Directory + '\' + ShortcutName + '.lnk';
   MyPFile.Save (PWChar (WFileName), False);
   MyReg.Free;
end;

function TirarAcentos( cString: String ): String;
var
   nPos: Integer;
begin
   for nPos := 1 to Length( cString ) do
   begin
      Case cString[nPos] of
         '¡', '¿', '¬', '√', 'ƒ' : cString[nPos] := 'A';
         '·', '‡', '‚', '„', '‰' : cString[nPos] := 'a';
         '«'                     : cString[nPos] := 'C';
         'Á'                     : cString[nPos] := 'c';
         '…', '»', ' ', 'À'      : cString[nPos] := 'E';
         'È', 'Ë', 'Í', 'Î'      : cString[nPos] := 'e';
         'Õ', 'Ã', 'Œ', 'œ'      : cString[nPos] := 'I';
         'Ì', 'Ï', 'Ó', 'Ô'      : cString[nPos] := 'i';
         '”', '“', '‘', '’', '÷' : cString[nPos] := 'O';
         'Û', 'Ú', 'Ù', 'ı', 'ˆ' : cString[nPos] := 'o';
         '⁄', 'Ÿ', '€', '‹'      : cString[nPos] := 'U';
         '˙', '˘', '˚', '¸'      : cString[nPos] := 'u';
      end;
   end;
   Result := cString;
end;

function TamanhoPapel(cTamPapel : string) : TQRPaperSize;
const
   aQRPaperSize : array[0..27] of string = ('Default', 'Letter', 'LetterSmall',
                     'Tabloid', 'Ledger', 'Legal', 'Statement', 'Executive',
                     'A3', 'A4', 'A4Small', 'A5', 'B4', 'B5', 'Folio',
                     'Quarto', 'qr10X14', 'qr11X17', 'Note', 'Env9', 'Env10',
                     'Env11', 'Env12', 'Env14', 'CSheet', 'DSheet', 'ESheet', 'Custom');
var
   nTamPap : integer;
begin
   nTamPap := BuscaVetorStr(aQRPaperSize, cTamPapel);

   case nTamPap of
     0 : Result := Default;
     1 : Result := Letter;
     2 : Result := LetterSmall;
     3 : Result := Tabloid;
     4 : Result := Ledger;
     5 : Result := Legal;
     6 : Result := Statement;
     7 : Result := Executive;
     8 : Result := A3;
     9 : Result := A4;
     10 : Result := A4Small;
     11 : Result := A5;
     12 : Result := B4;
     13 : Result := B5;
     14 : Result := Folio;
     15 : Result := Quarto;
     16 : Result := qr10X14;
     17 : Result := qr11X17;
     18 : Result := Note;
     19 : Result := Env9;
     20 : Result := Env10;
     21 : Result := Env11;
     22 : Result := Env12;
     23 : Result := Env14;
     24 : Result := CSheet;
     25 : Result := DSheet;
     26 : Result := ESheet;
     27 : Result := Custom;
   else ;
      Result := Default;
   end;
end;

function ChecaImpr : boolean;
begin
   Result := (Printer.Printers.Count > 0);
   if not Result then
      Application.MessageBox('Nenhuma impressora instalada !', 'AtenÁ„o',
         MB_OK+MB_ICONWARNING);
end;

function IndexImpr(cImpr : string) : integer;
var
   ImprDisp : TStringList;
   nPos : integer;
begin
   Result := Printer.PrinterIndex; // padr„o

   ImprDisp := TStringList.Create;
   ImprDisp.AddStrings(Printer.Printers);
   if ImprDisp.Count > 0 then
   begin
      nPos := ImprDisp.IndexOf(cImpr);
      if nPos >= 0 then
         Result := nPos;
   end;
   FreeAndNil(ImprDisp);
end;

function PesqTagNode(const Strtag: String; NodeIni : IXmlNode): IXmlNode;
{um objeto TXmlDocument cujo par‚metro do create for nil ou uma String
com o nome do Arquivo , devemos declarar o Tipo da variavel n„o com a
classe TXmlDocument e sim com a Interface IXmlDocument. }
var
   i, j : integer;
   XmlD : IXmlDocument; //Pela explicaÁÁ„o dada acima
   Node, NodeChild, NodeAux : IXmlNode;
begin
   Result := nil;
   if trim(Strtag) <> '' then
   try
      xmlD := TXmlDocument.Create(Application);
      try
         XmlD.Active := true;
         XmlD.DocumentElement := NodeIni;

         if XmlD.DocumentElement.NodeName = Trim(Strtag) then
            Result := XmlD.DocumentElement
         else
         for i := 0 to XmlD.DocumentElement.ChildNodes.Count - 1 do
         begin
            Node := XmlD.DocumentElement.ChildNodes[I];
            if Node.NodeName = Trim(Strtag) then
               Result := Node
            else
            for j := 0 to Node.ChildNodes.Count - 1 do
            begin
               NodeChild := Node.ChildNodes[j];
               if NodeChild.NodeName = Trim(Strtag) then
                  Result := NodeChild
               else
               begin
                  NodeAux := NodeChild.ChildNodes.FindNode(Strtag);
                  if Assigned(NodeAux) then
                     Result := NodeAux;
               end;
            end;
         end;
      except
       on e:exception do
       begin
        Showmessage('Problemas ao ler o Tag'+#13+
                     e.Message);
        Abort;
       end;
     end;
   finally
       // XmlD.Free; por ser Interface
   end;
end;

function CopyFiles(FromPath,ToPath,FileMask: String): Boolean;
var
   CopyFilesSearchRec: TSearchRec;
   FindFirstReturn: Integer;
   LastFile: string;
begin
   Result := False;
   FindFirstReturn :=
   FindFirst(FromPath+'\'+FileMask, faAnyFile, CopyFilesSearchRec);
   LastFile := CopyFilesSearchRec.Name;
   if not (CopyFilesSearchRec.Name = '') and
            not (FindFirstReturn = -18) then
   begin
      Result := True;
      CopyFile(PChar(FromPath+'\'+CopyFilesSearchRec.Name),PChar(ToPath+'\'+CopyFilesSearchRec.Name), True);
      while True Do
      begin
         if (FindNext(CopyFilesSearchRec)<0) or (LastFile = CopyFilesSearchRec.Name) Then
         begin
            Break;
         end
         else
         begin
            CopyFile(PChar(FromPath+'\'+CopyFilesSearchRec.Name),PChar(ToPath+'\'+CopyFilesSearchRec.Name), True);
            LastFile := CopyFilesSearchRec.Name;
         end;
      end;
   end;
end;

end.


