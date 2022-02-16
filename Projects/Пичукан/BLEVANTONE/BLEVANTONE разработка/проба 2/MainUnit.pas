unit MainUnit;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs, FMX.StdCtrls;

type
  TForm1 = class(TForm)
    ButtonOnFirstPanel: TButton;
    Exit: TButton;
    procedure ButtonOnFirstPanelClick(Sender: TObject);
    procedure ExitClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;



var
  Form1: TForm1;

implementation

{$R *.fmx}

uses AboutUnit;

procedure TForm1.ButtonOnFirstPanelClick(Sender: TObject);
     begin

      Form2.Show;


     end;

procedure TForm1.ExitClick(Sender: TObject);

begin


    Form1.Close;
  { Form1.Hide;       }
    { exit:= true;    }
   { Application.Destroy;    }
  { Application.free;         }


end;



end.
