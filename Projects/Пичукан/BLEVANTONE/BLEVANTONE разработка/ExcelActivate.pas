unit G5CH2;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    OpenDialog1: TOpenDialog;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    Button7: TButton;
    Button8: TButton;
    Button9: TButton;
    Button10: TButton;
    Button11: TButton;
    Button12: TButton;
    Button13: TButton;
    Button14: TButton;
    Button15: TButton;
    Button16: TButton;
    Button17: TButton;
    Button18: TButton;
    Button19: TButton;
    procedure Button3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure Button10Click(Sender: TObject);
    procedure Button11Click(Sender: TObject);
    procedure Button12Click(Sender: TObject);
    procedure Button13Click(Sender: TObject);
    procedure Button15Click(Sender: TObject);
    procedure Button16Click(Sender: TObject);
    procedure Button17Click(Sender: TObject);
    procedure Button18Click(Sender: TObject);
    procedure Button14Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}
  uses ComObj;
  var W:variant;
  text_:array[1..6] of string = ('фрагмент1','фрагмент2','фрагмент3','фрагмент4','фрагмент5','фрагмент6');

procedure TForm1.Button3Click(Sender: TObject);
begin
W.ActiveDocument.Close;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
W:=CreateOleObject('Word.Application');
W.Visible:=true;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
W.Quit;
W:=UnAssigned;
end;

procedure TForm1.Button2Click(Sender: TObject);
 var dir_:string;
begin
getdir(0,dir_);
if not OpenDialog1.Execute then begin exit; chDir(dir_); end;
chDir(dir_);
W.Documents.Open(OpenDialog1.FileName);
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
W.ActiveDocument.Range(0, 0).Select;
end;

procedure TForm1.Button4Click(Sender: TObject);
begin
W.ActiveDocument.Range(1, 50).Select;
end;

procedure TForm1.Button5Click(Sender: TObject);
 var eee_:string;
begin
eee_:=W.Selection.Text;
messagebox(handle,pchar(eee_),'Чтение текста из выделенного фрагмента!',0);
end;

procedure TForm1.Button6Click(Sender: TObject);
begin
W.Selection.Start:=W.ActiveDocument.Characters.Count;
W.Selection.End  :=W.ActiveDocument.Characters.Count;
end;

procedure TForm1.Button7Click(Sender: TObject);
begin
W.Selection.Text:='<-- вставляем текст -->';
end;

procedure TForm1.Button8Click(Sender: TObject);
begin
W.Selection.copy;
end;

procedure TForm1.Button9Click(Sender: TObject);
begin
W.Selection.paste;
end;

procedure TForm1.Button10Click(Sender: TObject);
begin
//W.Selection.delete;
//W.Selection.cut;
W.ActiveDocument.Range(1,20).Delete;
end;

procedure TForm1.Button11Click(Sender: TObject);
 var text_:string;
begin
text_:='';
text_:=InputBox('Введите текст для поиска',text_,text_);
messagebox(handle,pchar(text_),'Внимание!',0);
W.Selection.Find.Forward:=true;
W.Selection.Find.Text:=text_;
if W.Selection.Find.Execute then messagebox(handle,'Поиск текста завершен успешно!','Внимание!',0);
//if W.Selection.Find.Execute then W.Selection.Text:='Фрагмент текста для замены';
end;


procedure TForm1.Button12Click(Sender: TObject);
 const wdReplaceAll=2;
       wdFindContinue=1;
  var
       a_:integer;
begin
// Создаеем новый документ по шаблону
W.documents.Add(ExtractFileDir(Application.ExeName)+'\Поиск и замена текста.dot');
messagebox(handle,'Создан новый документ по шаблону!','Поиск и замена текста!',0);

W.Selection.Find.Forward:=true;
W.Selection.Find.Replacement.Text:=' <-- фрагмент для замены текста --> ';
for a_:=1 to 6 do begin
    W.Selection.Find.Text:=text_[a_];
    if W.Selection.Find.Execute(Replace:=wdReplaceAll) then messagebox(handle,pchar('Поиск текста "'+text_[a_]+'" завершен успешно!'),'Внимание!',0);
    messagebox(handle,pchar('Заменен текст "'+text_[a_]+'"'),'Поиск и замена текста!',0);
    end;
end;


procedure TForm1.Button13Click(Sender: TObject);
 const wdReplaceAll=2;
begin
// Создаеем новый документ по шаблону
W.documents.Add(ExtractFileDir(Application.ExeName)+'\Шаблон простого конверта.dot');

messagebox(handle,'Шаблон почтового конверта создан!','Внимание!',0);

// Подставляем индекс
W.Selection.Find.Text:='###индекс&';
W.Selection.Find.Replacement.Text:='350049';
W.Selection.Find.Execute(Replace:=wdReplaceAll);

W.Selection.Find.Text:='###адрес&';
W.Selection.Find.Replacement.Text:='Краснодар, ул. Севастопольская, д. 3, кв. 123';
W.Selection.Find.Execute(Replace:=wdReplaceAll);

W.Selection.Find.Text:='###получатель&';
W.Selection.Find.Replacement.Text:='Иванов Иван Иванович';
W.Selection.Find.Execute(Replace:=wdReplaceAll);

// Обратный адрес
W.Selection.Find.Text:='###обратный индекс&';
W.Selection.Find.Replacement.Text:='198005';
W.Selection.Find.Execute(Replace:=wdReplaceAll);

W.Selection.Find.Text:='###обратный адрес&';
W.Selection.Find.Replacement.Text:='Санкт-Петербург, Измайловский пр., д. 29, кв. 111';
W.Selection.Find.Execute(Replace:=wdReplaceAll);

W.Selection.Find.Text:='###отправитель&';
W.Selection.Find.Replacement.Text:='Петрова Альбина Васильевна';
W.Selection.Find.Execute(Replace:=wdReplaceAll);
end;


procedure TForm1.Button15Click(Sender: TObject);
 const wdCharacter=1;
begin
W.Selection.Move(wdCharacter,3);   //перемещение через 3 символа
end;

procedure TForm1.Button16Click(Sender: TObject);
begin
W.Selection.CopyAsPicture;
end;

procedure TForm1.Button17Click(Sender: TObject);
 const wdTableFormatGrid2=17;
begin
W.Selection.ConvertToTable(Separator:=' ',NumRows:=5,NumColumns:=5, Format:=wdTableFormatGrid2);
end;

procedure TForm1.Button18Click(Sender: TObject);
begin
W.Selection.TypeText('<-- вставляем текст -->');
end;

Function FindAndInsert(FindText,ReplacementText:string):boolean;
 const wdReplaceAll=2;
begin
W.Selection.Find.Text:=FindText;
W.Selection.Find.Replacement.Text:=ReplacementText;
FindAndInsert:=W.Selection.Find.Execute(Replace:=wdReplaceAll);
End;

procedure TForm1.Button14Click(Sender: TObject);
begin
// Создаеем новый документ по шаблону
W.documents.Add(ExtractFileDir(Application.ExeName)+'\Шаблон платежного поручения.dot');

messagebox(handle,'Шаблон создан! Переходим к заполнению.','Внимание!',0);

// Подставляем тектс
FindAndInsert('###№ П.П.&','1');
FindAndInsert('###Дата&',datetostr(date));
FindAndInsert('###Вид платежа&','почтой');
FindAndInsert('###Сумма прописью&','Двести пятьдесят рублей сорок копеек');
FindAndInsert('###Сумма&','250,40');
FindAndInsert('###ИНН плательщика&','0000000000');
FindAndInsert('###КПП плательщика&','000000000011');
FindAndInsert('###Плательщик&','ЗАО Селена');
FindAndInsert('###Р/С плательщика&','00000000000000000000');
FindAndInsert('###БИК плательщика&','000000');
FindAndInsert('###К/С плательщика&','00000000000000000000');

FindAndInsert('###ИНН получателя&','1111111111');
FindAndInsert('###КПП получателя&','111111111100');
FindAndInsert('###БИК получателя&','111111');
FindAndInsert('###К/С получателя&','11111111111111111111');
FindAndInsert('###Р/С получателя&','11111111111111111111');
FindAndInsert('###Получатель&','ЗАО Комета');

FindAndInsert('###В.О.&','');
FindAndInsert('###Н.П.&','');
FindAndInsert('###Код&','');
FindAndInsert('###С.П.&','');
FindAndInsert('###О.П.&','');
FindAndInsert('###Р.П.&','');

FindAndInsert('#Н1&','');
FindAndInsert('#Н2&','');
FindAndInsert('#Н3&','');
FindAndInsert('#Н4&','');
FindAndInsert('#Н5&','');
FindAndInsert('#Н6&','');
FindAndInsert('#Н7&','');

FindAndInsert('###Назначение платежа& ','Оплата за поставку товара');

end;



end.


------------------------------------------------------------------------------------------------------------

1
2
3
4
5
6
7
8
9
10
11
12
13
14
15
16
17
18
19
20
21
22
23
24
25
26
27
uses
  ComObj;
 
procedure TForm1.Button3Click(Sender: TObject);
var 
  ExcelApp, ExcelWB, ExcelWS : Variant; //Если Delphi 5, то OleVariant.
  i : integer;
begin
  ExcelApp := CreateOleObject('Excel.Application') ;
  ExcelApp.Visible := true;
  ExcelWB := ExcelApp.Workbooks.Open('c:\Свободное число\XLS.xls');
  ExcelWS := ExcelWB.Worksheets[1];
  
  i := 1;
  while ExcelWS.Cells[i, 1].Text <> '' do Inc(i);
  ExcelWS.Cells[i, 1].Value := Edit1.Text;
  
  ExcelApp.DisplayAlerts := False; //Отключаем режим показа предупреждений.
  try
    ExcelWB.Save; //Или, если сохраняем в другой файл: ExcelWB.SaveAs('c:\Свободное число\XLS.xls');
  finally
    ExcelApp.DisplayAlerts := True; //Включаем режим показа предупреждений.
  end;
  
  //Если требуется, закрываем MS Excel.
  ExcelApp.Quit;
end;
