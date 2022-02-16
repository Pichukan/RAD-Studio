unit Unit2;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Unit1;

type
  TForm2 = class(TForm)
    CreateButton: TButton;
    DelButton: TButton;
    procedure CreateButtonClick(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;
  massive: array [1..10, 1..10] of TEdit;
  massive2: array [1..10, 1..10] of Real;
  size: Integer;

implementation
{$R *.dfm}



procedure TForm2.CreateButtonClick(Sender: TObject);
var I,J : Integer;
begin
size:=Form1.SE1.Value;
for I := 1 to size do
begin
for J := 1 to size do
begin
massive [I,J]:=TEdit.Create(Self);
massive[I,J].Parent:=Self;
massive[I,J].Width:=40;
massive[I,J].Left:=30+J*50;
massive[I,J].Top:=15+I*30;
end;
end;
Form2.Height:=massive[size,size].Top+100;
Form2.Width:=massive[size,size].Top+300;

end;




procedure TForm2.DelButtonClick(Sender: TObject);
var i,j : Integer;
begin
if size <> 0 then

for i:=1 to size do
for j:=1 to size do
massive[i,j].Destroy;

end;



procedure TForm2.FormCreate(Sender: TObject);
begin
size:=0;
end;

end.
