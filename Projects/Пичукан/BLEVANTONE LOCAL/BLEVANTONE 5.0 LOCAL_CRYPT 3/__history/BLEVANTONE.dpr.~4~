program BLEVANTONE;

uses
  Vcl.Forms,
  Controls,
  Dialogs,
  SysUtils,
  WordActivate in 'WORD activate\WordActivate.pas' {Form1},
  Check in 'CHECK\Check.pas' {Form2};

{$R *.res}




begin

  Form2 := TForm2.Create(Nil);
 // Form2.ShowModal;
   Form2.Show;
 // Application.Initialize;
//  Application.CreateForm(MainForm, TMainForm);

   ShowMessage('1');
  Application.Initialize;
   ShowMessage('2');
  Application.MainFormOnTaskbar := True;
   ShowMessage('3');
 // Application.CreateForm(TForm2, Form2);
  Application.CreateForm(TForm1, Form1);
 // Application.CreateForm(TForm2, Form2);
  ShowMessage('4');
 // Form2 := TForm2.Create(Nil);
 // Form2.ShowModal;
 // Form1.Show;
  Application.Run;
   ShowMessage('5');
end.
