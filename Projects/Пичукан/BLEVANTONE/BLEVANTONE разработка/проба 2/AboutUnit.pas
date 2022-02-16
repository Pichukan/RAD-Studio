unit AboutUnit;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs, FMX.StdCtrls;

type
  TForm2 = class(TForm)
    AboutButton1: TButton;
    Label1: TLabel;
    procedure AboutButton1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

{$R *.fmx}

uses MainUnit;

procedure TForm2.AboutButton1Click(Sender: TObject);
begin

     { Form2.Hide;      }
    Form1.Show;
  {  Form2.Hide;      }

end;

end.
