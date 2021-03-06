unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, System.Win.ComObj, Winapi.ActiveX, TlHelp32, ShellAPI,
  Vcl.Samples.Spin;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Label1: TLabel;
    SE1: TSpinEdit;
    Button2: TButton;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  XLS: OleVariant;

implementation

{$R *.dfm}

uses Unit2;

function GetProcessByEXE(exename: string): THandle;
var
  hSnapshoot: THandle;
  pe32: TProcessEntry32;
begin
  Result:= 0;
  hSnapshoot:= CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  if (hSnapshoot = 0) then Exit;
  pe32.dwSize:= SizeOf(TProcessEntry32);
  if (Process32First(hSnapshoot, pe32)) then
    repeat
      if (pe32.szExeFile = exename) then
      begin
        Result:= pe32.th32ProcessID;
        exit;
      end;
    until not Process32Next(hSnapshoot, pe32);
end;


function  ExistExcel: Boolean;
   var
   ID:TCLSID;
   Res: HRESULT;
begin
    Res:= CLSIDFromProgID('Excel.Application',ID);
     if Res=S_OK then
     Result:= True
     else
     Result:= False;
end;

function RunExcel: Boolean;
begin

  if GetProcessByEXE('EXCEL.EXE')<>0 then

  begin
  XLS:=GetActiveOLEObject('Excel.Application');
  Result:=true;
  end
  else
  Result:= false;
  end;

 function StartExcel: boolean;
 begin
    if ExistExcel then
    begin
    if RunExcel=false then
    XLS:=CreateOleObject('Excel.Application');
    result:=True;
    end
    else
    begin
    ShowMessage('?????????? ????? ?? ?????????? ?? ?????? ?????????');
     result:=False;
    end;
 end;


procedure TForm1.Button1Click(Sender: TObject);
begin
 { if ExistExcel then
                ShowMessage('?????? ?????????? ?? ?????? ??????????')
                else
                ShowMessage('?????? ??????????? ?? ?????? ??????????');
   if RunExcel then
                ShowMessage('?????? ??????? ?? ?????? ??????????')
                else
                ShowMessage('?????? ?? ??????? ?? ?????? ??????????');
   RunExcel;  }

   if StartExcel then
   begin
    ShowMessage('??????? ???????');
   end;

end;





procedure TForm1.Button2Click(Sender: TObject);
begin
 Form2.DelButton.Click;
 Form2.CreateButton.Click;
 Form2.ShowModal;
end;

end.

