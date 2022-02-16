program BLEVANTONE;

uses
  Vcl.Forms,
  Controls,
  Dialogs,
  SysUtils,
  WordActivate in '..\BLEVANTONE 7.0 further development 1\WORD activate\WordActivate.pas' {Form1};

{$R *.res}




begin




  Application.Initialize;

  Application.MainFormOnTaskbar := True;

  Application.CreateForm(TForm1, Form1);
  Form1.Caption:=  ProgramName;

  Application.Run;
end.
