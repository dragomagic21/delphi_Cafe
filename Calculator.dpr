program Calculator;

uses
  Forms,
  uMain in 'uMain.pas' {fMain},
  uSostav in 'uSostav.pas' {fSostav};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := '����������� ����';
  Application.CreateForm(TfMain, fMain);
  Application.Run;
end.
