unit uSostav;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, AdvOfficePager, AdvPanel, ExtCtrls, Grids, BaseGrid, AdvGrid,
  DBAdvGrid, DB, ADODB, DBCtrls, StdCtrls, AdvGlowButton, Mask, AdvSpin,
  DBAdvSp, AdvReflectionLabel, AdvOfficeImage, DBTables, AdvObj;

type
  TfSostav = class(TForm)
    pSostav_Header: TAdvPanel;
    StylerHeader: TAdvPanelStyler;
    Pager: TAdvOfficePager;
    pSostav: TAdvOfficePage;
    pSostav_Nav: TAdvPanel;
    StylerPanel: TAdvPanelStyler;
    sgSostav: TDBAdvGrid;
    TB_BI: TADOTable;
    DS_BI: TDataSource;
    bSostavAdd: TDBAdvGlowButton;
    bSostavEdit: TDBAdvGlowButton;
    Bevel2: TBevel;
    lBludoName: TLabel;
    bSostavCancel: TDBAdvGlowButton;
    bSostavOk: TDBAdvGlowButton;
    bSostavDel: TDBAdvGlowButton;
    eSostavIngr: TDBLookupComboBox;
    TB_BIID: TAutoIncField;
    TB_BIBLUDO: TIntegerField;
    TB_BIINGR: TIntegerField;
    lIngrCount: TLabel;
    eSostavCount: TDBAdvSpinEdit;
    Lgram: TLabel;
    TB_BILINGR: TStringField;
    TB_BICOUNT: TIntegerField;
    bBludoFind: TAdvGlowButton;
    LBludo: TAdvReflectionLabel;
    LProgram: TLabel;
    TB_BILBLUDO: TStringField;
    TB_BIINGRCOST: TFloatField;
    TB_BICOST: TFloatField;
    Q: TADOQuery;
    LSostavCost: TLabel;
    procedure bSostavOkBeforeAction(Sender: TObject;
      var DoAction: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bBludoFindClick(Sender: TObject);
    procedure TB_BICalcFields(DataSet: TDataSet);
    procedure bSostavOkAfterAction(Sender: TObject;
      var ShowException: Boolean);
  private
    { Private declarations }
  public
    ID: integer;
    { Public declarations }
    function GetBludoName(): string;
    function GetCurrCost(): string;
  end;

var
  fSostav: TfSostav;

implementation

uses uMain;

{$R *.dfm}

procedure TfSostav.bSostavOkBeforeAction(Sender: TObject; var DoAction: Boolean);
begin
 TB_BIBLUDO.Value := ID;
 DoAction := TRUE;
end;

procedure TfSostav.FormCreate(Sender: TObject);
begin
 sgSostav.ColumnByFieldName['LINGR'].Header := 'Ингридиент';
 sgSostav.ColumnByFieldName['COUNT'].Header := 'Количество';
 sgSostav.ColumnByFieldName['COST'].Header := 'Общая стоимость';
end;

procedure TfSostav.FormResize(Sender: TObject);
begin
 sgSostav.ColumnByFieldName['ID'].Width := 0;
// sgSostav.ColumnByFieldName['BLUDO'].Width := 0;
// sgSostav.ColumnByFieldName['INGR'].Width := 0;
 sgSostav.ColumnByFieldName['COUNT'].Width := 100;
 sgSostav.ColumnByFieldName['COST'].Width := 110;
 sgSostav.ColumnByFieldName['LINGR'].Width := (sgSostav.Width-250);
end;

procedure TfSostav.FormShow(Sender: TObject);
begin
 // Показываем состав только выбранного блюда
 TB_BI.Open;
 DS_BI.Enabled := TRUE;
 TB_BI.Filter := 'BLUDO = '+IntToStr(ID);
 TB_BI.Filtered := TRUE;
 lBludo.HTMLText.Text := '<P align="center"><FONT face="Comic Sans MS" size="14">"'+GetBludoName+'"</FONT></P>';
 Self.Caption := 'Калькулятор блюд. '+GetBludoName+'.';
 LSostavCost.Caption := 'Стоимость: '+GetCurrCost+' грн.';
end;

procedure TfSostav.bBludoFindClick(Sender: TObject);
begin
 sgSostav.SearchFooter.Visible := not(sgSostav.SearchFooter.Visible);
end;

function TfSostav.GetBludoName: string;
begin
 Result := '';
 Q.Close();
 Q.SQL.Clear;
 Q.SQL.Add('select TB_BLUDO.NAME from TB_BLUDO where TB_BLUDO.ID = '+IntToStr(ID));
 Q.Open;
 Result :=  Q.Fields[0].AsString;
 Q.Close();
end;

procedure TfSostav.TB_BICalcFields(DataSet: TDataSet);
begin
 TB_BICOST.Value := TB_BICOUNT.Value*TB_BIINGRCOST.Value;
end;

procedure TfSostav.bSostavOkAfterAction(Sender: TObject; var ShowException: Boolean);
var i, Col: integer;
    cost: real;
begin
 Col := sgSostav.ColumnByFieldName['COST'].Index;
 cost := 0;
 // Выбираем общую сумму
 for i := 1 to sgSostav.RowCount-1 do
  cost := cost + StrToFloat(sgSostav.Cells[Col,i]);

 // Выполняем запрос
 Q.Close();
 Q.SQL.Clear;
 Q.SQL.Add('UPDATE TB_BLUDO SET TB_BLUDO.COST = '''+FloatToStr(cost)+''' WHERE TB_BLUDO.ID = '+IntToStr(ID)+';');
 Q.ExecSQL;
 Q.Close();
 fMain.TB_BLUDO.Refresh;
end;

function TfSostav.GetCurrCost: string;
var i: integer;
    sum: real;
begin
 Result := '0.00';
 if (sgSostav.Cells[1,1] = '') then exit;
 sum := 0;
 for i := 1 to sgSostav.RowCount-1 do
  sum := sum + StrToFloat(sgSostav.Cells[sgSostav.ColumnByFieldName['COST'].Index,i]);
 Result := FloatToStr(sum);
end;

end.
