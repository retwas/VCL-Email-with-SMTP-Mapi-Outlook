program Email;

uses
  Vcl.Forms,
  uFormEmail in 'uFormEmail.pas' {FormEmail};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFormEmail, FormEmail);
  Application.Run;
end.
