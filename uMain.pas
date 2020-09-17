unit uMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, DBTables, Grids, BaseGrid, AdvGrid, DBAdvGrid, ExtCtrls,
  AdvPanel, AdvOfficePager, StdCtrls, DBCtrls, AdvGlowButton, ComCtrls,
  AdvReflectionLabel, AdvPicture, Mask, AdvEdit, AdvEdBtn,
  PlannerDatePicker, PlannerDBDatePicker, Buttons, IniFiles, AdvSpin,
  DBAdvSp, OleServer, WordXP, ADODB, MoneyEdit, dbmnyed, AdvOfficeImage,
  ShellApi, AdvSmoothPanel, AdvSmoothExpanderPanel, ToolPanels, AdvObj;

type
  TfMain = class(TForm)
    Pager: TAdvOfficePager;
    StylerHeader: TAdvPanelStyler;
    StylerPanel: TAdvPanelStyler;
    pBludo: TAdvOfficePage;
    BaseDB: TADOConnection;
    TB_BLUDO: TADOTable;
    TB_BLUDOID: TAutoIncField;
    TB_BLUDONAME: TWideStringField;
    DS_BLUDO: TDataSource;
    DS_INGR: TDataSource;
    TB_INGR: TADOTable;
    pBludo_Nav: TAdvPanel;
    sgBludo: TDBAdvGrid;
    TB_INGRID: TAutoIncField;
    TB_INGRNAME: TWideStringField;
    TB_INGRCOST: TBCDField;
    pIngr: TAdvOfficePage;
    sgIngr: TDBAdvGrid;
    pIngr_Nav: TAdvPanel;
    bIngrAdd: TDBAdvGlowButton;
    bIngrPost: TDBAdvGlowButton;
    bIngrCancel: TDBAdvGlowButton;
    bIngrEdit: TDBAdvGlowButton;
    bIngrDel: TDBAdvGlowButton;
    eIngrName: TDBEdit;
    lIngrName: TLabel;
    lIngrCost: TLabel;
    eIngrCost: TDBMoneyEdit;
    Bevel1: TBevel;
    bIngrFind: TAdvGlowButton;
    bBludoAdd: TDBAdvGlowButton;
    bBludoEdit: TDBAdvGlowButton;
    Bevel2: TBevel;
    lBludoName: TLabel;
    eBludoName: TDBEdit;
    bBludoCancel: TDBAdvGlowButton;
    bBludoOk: TDBAdvGlowButton;
    bBludoDel: TDBAdvGlowButton;
    bBludoFind: TAdvGlowButton;
    Q_BLUDO_COST: TADOQuery;
    DS_BLUDO_COST: TDataSource;
    TB_BLUDOCOST: TFloatField;
    bBludoSostav: TAdvGlowButton;
    eBludoSort: TRadioGroup;
    eIngrSort: TRadioGroup;
    WordA: TWordApplication;
    bPrint: TAdvGlowButton;
    AdvGlowButton1: TAdvGlowButton;
    Bevel3: TBevel;
    Bevel4: TBevel;
    Bevel5: TBevel;
    pSklad: TAdvOfficePage;
    Pager_Sklad: TAdvOfficePager;
    pSklad_All: TAdvOfficePage;
    pSklad_Pr: TAdvOfficePage;
    pSklad_RS: TAdvOfficePage;
    pKoef: TAdvOfficePage;
    sgKoef: TDBAdvGrid;
    pKoef_Nav: TAdvPanel;
    TB_KEF: TADOTable;
    TB_KEFID: TAutoIncField;
    TB_KEFIN: TIntegerField;
    TB_KEFOUT: TIntegerField;
    TB_KEFLOUT: TStringField;
    DS_KEF: TDataSource;
    bKoefAdd: TDBAdvGlowButton;
    bKoefEdit: TDBAdvGlowButton;
    Bevel6: TBevel;
    bKoefCancel: TDBAdvGlowButton;
    bKoefOk: TDBAdvGlowButton;
    bKoefDel: TDBAdvGlowButton;
    LKoef_IN: TLabel;
    LKoef_OUT: TLabel;
    LKoef: TLabel;
    eKoef: TDBAdvSpinEdit;
    LKoef_L: TLabel;
    eKoefIN: TDBLookupListBox;
    eKoefOUT: TDBLookupListBox;
    TB_KEFINO: TStringField;
    sgSKLAD: TDBAdvGrid;
    AdvPanel1: TAdvPanel;
    TB_SKLAD: TADOTable;
    DS_SKLAD: TDataSource;
    TB_SKLADID: TAutoIncField;
    TB_SKLADINGR: TIntegerField;
    TB_SKLADCOUNT: TIntegerField;
    TB_SKLADLINGR: TStringField;
    DS_PR: TDataSource;
    TB_PR: TADOTable;
    pSkladAdd: TAdvPanel;
    bSkladAddCancel: TDBAdvGlowButton;
    bSkladAddOk: TDBAdvGlowButton;
    eSkladAddPrim: TDBEdit;
    LSkladAddPrim: TLabel;
    eSkladAddDate: TPlannerDBDatePicker;
    LSkladAddDate: TLabel;
    eSkladAddCount: TDBAdvSpinEdit;
    LSkladAddCount: TLabel;
    LSkladAddL: TLabel;
    eSkladAddIngr: TDBLookupListBox;
    LSkladAddIngr: TLabel;
    bSkladAdd: TDBAdvGlowButton;
    Bevel7: TBevel;
    pSkladDel: TAdvPanel;
    Bevel8: TBevel;
    LSkladDelPrim: TLabel;
    LSkladDelDate: TLabel;
    LSkladDelCount: TLabel;
    LSkladDelL: TLabel;
    LSkladDelIngr: TLabel;
    bSkladCancel: TDBAdvGlowButton;
    bSkladOk: TDBAdvGlowButton;
    eSkladDelPrim: TDBEdit;
    eSkladDelDate: TPlannerDBDatePicker;
    eSkladDelCount: TDBAdvSpinEdit;
    eSkladDelIngr: TDBLookupListBox;
    bSkladDel: TDBAdvGlowButton;
    bSkladPrint: TAdvGlowButton;
    bSkladFind: TAdvGlowButton;
    sgPR: TDBAdvGrid;
    sgRS: TDBAdvGrid;
    TB_PRID: TAutoIncField;
    TB_PRINGR: TIntegerField;
    TB_PRCOUNT: TIntegerField;
    TB_PRDATE: TDateTimeField;
    TB_PRPRIM: TWideStringField;
    TB_PRLINGR: TStringField;
    AdvGlowButton2: TAdvGlowButton;
    Q: TADOQuery;
    TB_RS: TADOTable;
    DS_RS: TDataSource;
    TB_KEFCOEF: TFloatField;
    TB_RSID: TAutoIncField;
    TB_RSINGR: TIntegerField;
    TB_RSCOUNT: TIntegerField;
    TB_RSDATE: TDateTimeField;
    TB_RSPRIM: TWideStringField;
    TB_RSStringField: TStringField;
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure bIngrAddClick(Sender: TObject);
    procedure bIngrEditClick(Sender: TObject);
    procedure bIngrFindClick(Sender: TObject);
    procedure bBludoFindClick(Sender: TObject);
    procedure bBludoAddClick(Sender: TObject);
    procedure bBludoEditClick(Sender: TObject);
    procedure bBludoOkBeforeAction(Sender: TObject; var DoAction: Boolean);
    procedure bBludoSostavClick(Sender: TObject);
    procedure eBludoSortClick(Sender: TObject);
    procedure eIngrSortClick(Sender: TObject);
    procedure LemailClick(Sender: TObject);
    procedure bPrintClick(Sender: TObject);
    procedure AdvGlowButton1Click(Sender: TObject);
    procedure bKoefAddClick(Sender: TObject);
    procedure bKoefEditClick(Sender: TObject);
    procedure bSkladAddClick(Sender: TObject);
    procedure bSkladDelClick(Sender: TObject);
    procedure bSkladFindClick(Sender: TObject);
    procedure AdvGlowButton2Click(Sender: TObject);
    procedure bSkladAddOkAfterAction(Sender: TObject;
      var ShowException: Boolean);
    procedure bSkladOkBeforeAction(Sender: TObject; var DoAction: Boolean);
  private
    { Private declarations }
  public
    { Public declarations }
    f: TIniFile;
  end;

var
  fMain: TfMain;

implementation

uses Math, uSostav;

{$R *.dfm}

procedure TfMain.FormCreate(Sender: TObject);
begin
 Pager.ActivePageIndex := 0;
 
 // даем имена стобцам гридов
 sgBludo.ColumnByFieldName['NAME'].Header := 'Название';
 sgBludo.ColumnByFieldName['COST'].Header := 'Стоимость';

 sgIngr.ColumnByFieldName['NAME'].Header := 'Название';
 sgIngr.ColumnByFieldName['COST'].Header := 'Стоимость';

 sgKoef.ColumnByFieldName['LIN'].Header := 'Вошло';
 sgKoef.ColumnByFieldName['LOUT'].Header := 'Вышло';
 sgKoef.ColumnByFieldName['COEF'].Header := 'Коэфициент';

 sgSKLAD.ColumnByFieldName['LINGR'].Header := 'Ингридиент';
 sgSKLAD.ColumnByFieldName['COUNT'].Header := 'Кол-во грамм';

 sgPR.ColumnByFieldName['LINGR'].Header := 'Ингридиент';
 sgPR.ColumnByFieldName['COUNT'].Header := 'Кол-во грамм';
 sgPR.ColumnByFieldName['DATE'].Header := 'Дата';
 sgPR.ColumnByFieldName['PRIM'].Header := 'Примечание';

 sgRS.ColumnByFieldName['LINGR'].Header := 'Ингридиент';
 sgRS.ColumnByFieldName['COUNT'].Header := 'Кол-во грамм';
 sgRS.ColumnByFieldName['DATE'].Header := 'Дата';
 sgRS.ColumnByFieldName['PRIM'].Header := 'Примечание';

 eSkladAddDate.Date := Now();
 eSkladDelDate.Date := Now();

 // открываем нашу БД
 try
  f := TIniFile.Create(GetCurrentDir+'\config.ini');
  BaseDB.ConnectionString := f.ReadString('DATABASE','base',
  'Provider=Microsoft.Jet.OLEDB.4.0;Password="";User ID=Admin;Data Source=base.mdb;Mode=ReadWrite;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Database Password="";'+
  'Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;'+
  'Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don''t Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False');
  BaseDB.Open;
  TB_BLUDO.Open;
  DS_BLUDO.Enabled := TRUE;
  TB_INGR.Open;
  DS_INGR.Enabled := TRUE;
  TB_KEF.Open;
  DS_KEF.Enabled := TRUE;
  TB_SKLAD.Open;
  DS_SKLAD.Enabled := TRUE;
  TB_PR.Open;
  DS_PR.Enabled := TRUE;
  TB_RS.Open;
  DS_RS.Enabled := TRUE;
 except
 // MessageBox(Self.Handle,'Произошла ошибка при подключении к базе данных. Проверьте настройки файла конфигурации config.ini!','Ошибка!',MB_OK or MB_ICONWARNING);
 end;
end;


procedure TfMain.FormResize(Sender: TObject);
begin
 sgBludo.ColumnByFieldName['ID'].Width := 0;
 sgBludo.ColumnByFieldName['NAME'].Width := (sgIngr.Width-140);
 sgBludo.ColumnByFieldName['COST'].Width := 100;

 sgIngr.ColumnByFieldName['ID'].Width := 0;
 sgIngr.ColumnByFieldName['NAME'].Width := (sgIngr.Width-140);
 sgIngr.ColumnByFieldName['COST'].Width := 100;

 sgKoef.ColumnByFieldName['ID'].Width := 0;
 sgKoef.ColumnByFieldName['COEF'].Width := 100;
 sgKoef.ColumnByFieldName['LIN'].Width := (sgIngr.Width-140) div 2;
 sgKoef.ColumnByFieldName['LOUT'].Width := (sgIngr.Width-140) div 2;

 sgSKLAD.ColumnByFieldName['ID'].Width := 0;
 sgSKLAD.ColumnByFieldName['LINGR'].Width := (sgSKLAD.Width-140);
 sgSKLAD.ColumnByFieldName['COUNT'].Width := 100;

 sgPR.ColumnByFieldName['ID'].Width := 0;
 sgPR.ColumnByFieldName['LINGR'].Width := (sgPR.Width-240) div 2;
 sgPR.ColumnByFieldName['COUNT'].Width := 100;
 sgPR.ColumnByFieldName['DATE'].Width := 100;
 sgPR.ColumnByFieldName['PRIM'].Width := (sgPR.Width-240) div 2;

 sgRS.ColumnByFieldName['ID'].Width := 0;
 sgRS.ColumnByFieldName['LINGR'].Width := (sgRS.Width-240) div 2;
 sgRS.ColumnByFieldName['COUNT'].Width := 100;
 sgRS.ColumnByFieldName['DATE'].Width := 100;
 sgRS.ColumnByFieldName['PRIM'].Width := (sgRS.Width-240) div 2;
end;

procedure TfMain.bIngrAddClick(Sender: TObject);
begin
 eIngrName.SetFocus;
end;

procedure TfMain.bIngrEditClick(Sender: TObject);
begin
 eIngrName.SetFocus;
end;

procedure TfMain.bIngrFindClick(Sender: TObject);
begin
 sgIngr.SearchFooter.Visible := not(sgIngr.SearchFooter.Visible);
end;

procedure TfMain.bBludoFindClick(Sender: TObject);
begin
 sgBludo.SearchFooter.Visible := not(sgBludo.SearchFooter.Visible);
end;

procedure TfMain.bBludoAddClick(Sender: TObject);
begin
 eBludoName.SetFocus;
end;

procedure TfMain.bBludoEditClick(Sender: TObject);
begin
 eBludoName.SetFocus;
end;

procedure TfMain.bBludoOkBeforeAction(Sender: TObject; var DoAction: Boolean);
begin
 TB_BLUDOCOST.Value := 0.00;
 DoAction := TRUE;
end;

procedure TfMain.bBludoSostavClick(Sender: TObject);
var Col, Row: integer;
begin
 Col := sgBludo.ColumnByFieldName['ID'].Index;
 Row := sgBludo.Row;
 Application.CreateForm(TfSostav, fSostav);
 fSostav.ID := StrToInt(sgBludo.Cells[Col,Row]);
 fSostav.ShowModal;
end;

procedure TfMain.eBludoSortClick(Sender: TObject);
begin
 case eBludoSort.ItemIndex of
  0: TB_BLUDO.IndexFieldNames := 'NAME';
  1: TB_BLUDO.IndexFieldNames := 'COST';
 end;
end;

procedure TfMain.eIngrSortClick(Sender: TObject);
begin
 case eIngrSort.ItemIndex of
  0: TB_INGR.IndexFieldNames := 'NAME';
  1: TB_INGR.IndexFieldNames := 'COST';
 end;
end;

procedure TfMain.LemailClick(Sender: TObject);
begin
 ShellExecute(Handle, nil, 'mailto:drago_magic@mail.ru', nil, nil, SW_SHOW);
end;

procedure TfMain.bPrintClick(Sender: TObject);
var FileName: OleVariant;
    i, j: integer;
    Wt: Table;
begin
 FileName:=GetCurrentDir+'\Report.dot';
// with WordA do
 try  // Word ?? ???????, ?????????
  WordA.Disconnect;
  WordA.Connect;
  WordA.Visible := TRUE;
  WordA.Documents.OpenOld(FileName,EmptyParam,EmptyParam,EmptyParam,
                          EmptyParam,EmptyParam,EmptyParam,
	                  EmptyParam,EmptyParam,EmptyParam);
  SelectFirst;
  WordA.Selection.NextField;
  while (WordA.Selection.Text <> 'q')or(WordA.Selection.Text <> 'Q') do
  begin
   case WordA.Selection.Text[1] of
    'q','Q': break;

    'n','N': WordA.Selection.Text := sgBludo.Cells[sgBludo.ColumnByFieldName['NAME'].Index, sgBludo.Row];

    't','T': begin
              WordA.Selection.Text := '';
              Application.CreateForm(TfSostav, fSostav);
              fSostav.ID := StrToInt(sgBludo.Cells[sgBludo.ColumnByFieldName['ID'].Index,sgBludo.Row]);
              fSostav.Show;
              Wt := WordA.ActiveDocument.Tables.AddOld(WordA.Selection.Range,fSostav.sgSostav.RowCount,fSostav.sgSostav.ColCount-2);
              for i := 2 to fSostav.sgSostav.ColCount-1 do
               for j := 0 to fSostav.sgSostav.RowCount-1 do
                Wt.Cell(j+1,i-1).Range.Text := fSostav.sgSostav.Cells[i,j];
//              WordA.Selection.NextField();
              fSostav.Close();
             end;
    'i','I': WordA.Selection.Text := 'Итого: '+sgBludo.Cells[sgBludo.ColumnByFieldName['COST'].Index, sgBludo.Row]+' грн.';

   end;
   WordA.Selection.NextField();
  end;
  WordA.Selection.Text := '';
 except
  WordA.Disconnect;
  MessageBox(Self.Handle,'Ошибка! Не удается найти Microsoft Word. Установка этого приложения исправит проблему.','Ошибка!',MB_ICONWARNING or MB_OK);
 end;
end;

procedure TfMain.AdvGlowButton1Click(Sender: TObject);
var FileName: OleVariant;
    i, j, bl: integer;
    Wt: Table;
begin
 FileName:=GetCurrentDir+'\Report.dot';
// with WordA do
 try  // Word ?? ???????, ?????????
  bl := 1;
  WordA.Disconnect;
  WordA.Connect;
  WordA.Visible := TRUE;
  WordA.Documents.OpenOld(FileName,EmptyParam,EmptyParam,EmptyParam,
                          EmptyParam,EmptyParam,EmptyParam,
	                  EmptyParam,EmptyParam,EmptyParam);
  SelectFirst;
  WordA.Selection.NextField;
  while (WordA.Selection.Text <> 'q')or(WordA.Selection.Text <> 'Q') do
  begin
   case WordA.Selection.Text[1] of
    'q','Q': break;

    'n','N': WordA.Selection.Text := sgBludo.Cells[sgBludo.ColumnByFieldName['NAME'].Index, sgBludo.Row];

    't','T': begin
              WordA.Selection.Text := '';
              for bl := 1 to sgBludo.RowCount-1 do
              begin
               Application.CreateForm(TfSostav, fSostav);
               fSostav.ID := StrToInt(sgBludo.Cells[sgBludo.ColumnByFieldName['ID'].Index,sgBludo.Row]);
               fSostav.Show;
               Wt := WordA.ActiveDocument.Tables.AddOld(WordA.Selection.Range,fSostav.sgSostav.RowCount,fSostav.sgSostav.ColCount-2);
               for i := 2 to fSostav.sgSostav.ColCount-1 do
                for j := 0 to fSostav.sgSostav.RowCount-1 do
                 Wt.Cell(j+1,i-1).Range.Text := fSostav.sgSostav.Cells[i,j];
//             ВСТАВИТЬ ПЕРЕНОС СТРОКИ: ^p - И ВСЕ - ОСТАЛЬНОЕ РАБОТАЕТ
               fSostav.Close();
//               WordA.ActiveDocument.Paragraphs.Add()
              end;
             WordA.Selection.NextField();
             WordA.Selection.NextField();  
             end;
//    'i','I': WordA.Selection.Text := 'Итого: '+sgBludo.Cells[sgBludo.ColumnByFieldName['COST'].Index, sgBludo.Row]+' грн.';

   end;
   WordA.Selection.NextField();
  end;
  WordA.Selection.Text := '';
 except
  WordA.Disconnect;
  MessageBox(Self.Handle,'Ошибка! Не удается найти Microsoft Word. Установка этого приложения исправит проблему.','Ошибка!',MB_ICONWARNING or MB_OK);
 end;
end;

procedure TfMain.bKoefAddClick(Sender: TObject);
begin
 eKoefIN.SetFocus;
end;

procedure TfMain.bKoefEditClick(Sender: TObject);
begin
 eKoefIN.SetFocus;
end;

procedure TfMain.bSkladAddClick(Sender: TObject);
begin
 eSkladAddIngr.SetFocus();
end;

procedure TfMain.bSkladDelClick(Sender: TObject);
begin
 eSkladDelIngr.SetFocus();
end;

procedure TfMain.bSkladFindClick(Sender: TObject);
begin
 sgSKLAD.SearchFooter.Visible := not(sgSKLAD.SearchFooter.Visible);
end;

procedure TfMain.AdvGlowButton2Click(Sender: TObject);
begin
 sgPR.SearchFooter.Visible := not(sgPR.SearchFooter.Visible);
end;

procedure TfMain.bSkladAddOkAfterAction(Sender: TObject; var ShowException: Boolean);
var i: byte;
    s: integer;
begin
 // смотрим есть ли ингридиент на складе
 Q.Close();
 Q.SQL.Clear;
 Q.SQL.Add('SELECT COUNT(TB_SKLAD.ID) FROM TB_SKLAD WHERE TB_SKLAD.INGR= '+IntToStr(integer(eSkladAddIngr.KeyValue))+';');
 Q.Open;
 i := Q.Fields[0].AsInteger;
 Q.Close();

 // если 0 - insert (нет такого), else update
 case i of
  0: begin
       Q.Close();
       Q.SQL.Clear;
       Q.SQL.Add('INSERT INTO TB_SKLAD (INGR, [COUNT]) VALUES ('+IntToStr(integer(eSkladAddIngr.KeyValue))+', '+IntToStr(eSkladAddCount.Value)+');');
       Q.ExecSQL;
       Q.Close();
       TB_SKLAD.Refresh;
     end;
  1: begin
      // Тырим существующее кол-во
      Q.Close();
      Q.SQL.Clear;
      Q.SQL.Add('SELECT TB_SKLAD.COUNT FROM TB_SKLAD WHERE TB_SKLAD.INGR= '+IntToStr(integer(eSkladAddIngr.KeyValue))+';');
      Q.Open;
      s := Q.Fields[0].AsInteger;
      Q.Close();

      Q.Close();
      Q.SQL.Clear;
      Q.SQL.Add('UPDATE TB_SKLAD SET TB_SKLAD.COUNT = '+FloatToStr(s+eSkladAddCount.Value)+' WHERE TB_SKLAD.INGR = '+IntToStr(integer(eSkladAddIngr.KeyValue))+';');
      Q.ExecSQL;
      Q.Close();
      TB_SKLAD.Refresh;
     end;
 end;

 sgSKLAD.Refresh;
end;

procedure TfMain.bSkladOkBeforeAction(Sender: TObject; var DoAction: Boolean);
var list: TListBox;
    i, z: integer;
    kef: real;
begin
{
 DoAction := False;
 // Выбираем все ингридиеты блюда
 Q.Close();
 Q.SQL.Clear;
 Q.SQL.Add('SELECT TB_BI.INGR FROM TB_BI WHERE TB_BI.BLUDO= '+IntToStr(integer(eSkladDelIngr.KeyValue))+';');
 Q.Open;
 list := TListBox.Create(self);
 while not(Q.Eof) do
 begin
  list.Items.Add(Q.Fields[0].AsString);
  Q.Next;
 end;
 Q.Close();

 // бежим по каждому ингридиенту и смотрим есть ли он в коефициентах
 for i:= 0 to list.Items.Count-1 do
 begin
  Q.Close();
  Q.SQL.Clear;
  Q.SQL.Add('SELECT COUNT(TB_KEF.ID) FROM TB_KEF WHERE TB_KEF.OUT= '+list.Items[i]+';');
  Q.Open;
  z := Q.Fields[0].AsInteger;
  Q.Close();

  // если есть, то херачим по коефициенту
  if (z = 1) then
  // тырим коэфициент
  begin
   Q.Close();
   Q.SQL.Clear;
   Q.SQL.Add('SELECT TB_KEF.COEF FROM TB_KEF WHERE TB_KEF.OUT= '+list.Items[i]+';');
   Q.Open;
   kef := Q.Fields[0].AsFloat;
   Q.Close();

   // добавляем в расходы
   Q.Close();
   Q.SQL.Clear;
   Q.SQL.Add('INSERT INTO TB_SKLAD (INGR, [COUNT], [DATE], [PRIM]) VALUES ('+list.Items[i]+', '+IntToStr(eSkladAddCount.Value)+');');
   Q.ExecSQL;
   Q.Close();
  end;

 end;

 list.Free;
 }
end;

end.
