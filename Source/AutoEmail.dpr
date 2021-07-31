program AutoEmail;

uses
  Vcl.Forms,
  UnitAutoEMail in 'UnitAutoEMail.pas' {FormMain};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := False;
  Application.CreateForm(TFormMain, FormMain);
  Application.Run;
end.
