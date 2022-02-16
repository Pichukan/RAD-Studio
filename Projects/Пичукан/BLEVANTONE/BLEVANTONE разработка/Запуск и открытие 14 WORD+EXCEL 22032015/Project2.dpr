program Project2;

uses
  Vcl.Forms,
  WordActivate in 'WORD activate\WordActivate.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
