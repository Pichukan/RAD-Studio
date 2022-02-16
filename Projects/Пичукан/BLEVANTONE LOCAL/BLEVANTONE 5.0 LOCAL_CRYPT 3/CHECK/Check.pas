unit Check;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls;

type
  TForm2 = class(TForm)
    Label1: TLabel;
    Edit1: TEdit;
    Button1: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

{$R *.dfm}





procedure TForm2.Button1Click(Sender: TObject);
begin
 // Form2.Show;
 //Form2.Hide;
 // Form2.Close;
 //Form2.ModalResult:=0;
  //TForm2.FormClose(Sender: TObject; var Action: TCloseAction);
end;

procedure TForm2.FormClose(Sender: TObject; var Action: TCloseAction);
begin
ShowMessage('Close me');
 // Form2.Close;
 // Form2.Free;

end;


end.
