program Project1;

uses
  System.StartUpCopy,
  FMX.Forms,
  MainUnit in 'MainUnit.pas' {Form1},
  AboutUnit in 'AboutUnit.pas' {Form2};

{$R *.res}
     var
      exit:boolean;
begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TForm2, Form2);
  Application.Run;

end.
