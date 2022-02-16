unit WordActivate;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, System.Win.ComObj, Winapi.ActiveX, TlHelp32, ShellAPI,
  Vcl.Samples.Spin;

type
  TForm1 = class(TForm)
    Button2: TButton;
    OpenDialog1: TOpenDialog;


    procedure Button2Click(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  WRD,Book,BookPicture,BookObzor,BookAnalog,BookLocation,
      BookBlevantone,RangeBookBlevantone,wdInlineShapes,wdInlineShapes2: OleVariant;
  EXC,MyBook,MyWorkSheet,MyWorkSheet2,MyRange,MyRange2,RangeObzor,
      Shp,vstart,vend: OleVariant;
  var W,ObzorValue:variant;
                 i:Integer;
        LengthDir : Integer;
  var DIRFName: string;
  var DIRName: string;
  var DIRExName: string;
  var NewWordDocDir: string;
  var FileFormat: OleVariant;
  var ProgramName: string;
  var  DIRFile: string;

implementation

{$R *.dfm}

//*************** проверка наличия действующего процесса EXCEL ****************
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

//***************** проверка установлен ли EXCEL ************************
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

//************** проверка установлен ли WORD **********************************
function  ExistWORD: Boolean;
   var
   ID:TCLSID;
   Res: HRESULT;
begin
    Res:= CLSIDFromProgID('Word.Application',ID);
     if Res=S_OK then
     Result:= True
     else
     Result:= False;
end;

//************* функция запуска приложения из запущенного EXCEL ***********************************

function RunExcel: Boolean;
begin

  if GetProcessByEXE('Excel.EXE')<>0 then

  begin
  EXC:=GetActiveOLEObject('Excel.Application');
  Result:=true;
  end
  else
  Result:= false;
  end;


//*********************** функция запуска приложения из запущенного WORD *******************************
function RunWORD: Boolean;
begin

  if GetProcessByEXE('WORD.EXE')<>0 then

  begin
  WRD:=GetActiveOLEObject('Word.Application');
  Result:=true;
  end
  else
  Result:= false;
  end;

//*********************** функция запуска приложения EXCEL *********************
  function StartExcel: boolean;
 begin
    if ExistExcel then
    begin
    if RunExcel=false then
    EXC:=CreateOleObject('Excel.Application');
    result:=True;
    end
    else
    begin
    ShowMessage('Майкрософт Excel не установлен на данном компютере');
     result:=False;
    end;
 end;


//*********************** функция запуска приложения WORD *******************************
 function StartWORD: boolean;
 begin
    if ExistWORD then
    begin
    if RunWORD=false then
    WRD:=CreateOleObject('Word.Application');
    result:=True;
    end
    else
    begin
    ShowMessage('Майкрософт WORD не установлен на данном компютере');
     result:=False;
    end;
 end;

 //**** определяется директория с названием файла без расширения  но с точкой ****************
 function DIRDetect: string;

 var

 Smb      : Char;
 i        : Integer;
 DIRECTORY: string;

 begin
       DIRECTORY:=' ';
       i:=2;
        DIRECTORY:=DIRFname[1];

        while (Smb<>'.') do
          begin
          Smb:=DIRFname[i];
          DIRECTORY:=DIRECTORY+Smb;
          i:=i+1;
          end;

       result:=DIRECTORY;

 end;

//****************** определяется директория папки без названия файла **************
 function DIRFileDetect: string;

 var

 Smb         : Char;
 i , L       : Integer;
 DIRECTORY   : string;

 begin
        DIRECTORY:=' ';
        DIRECTORY:=DIRFname;
        L:=Length(DIRFname);
        i:=L;
        Smb:=DIRFname[i];

        while (Smb<>'\') do
          begin
             i:=i-1;
             Smb:=DIRFname[i];
          end;

          Delete(DIRECTORY, i+1, l-i);
          LengthDir:=  Length(DIRECTORY);

          result:=DIRECTORY;

 end;


 // ***************** функция ищет в ВОРДЕ заданный текст по всему документу ************************************************
 // *****************         и дает на него ссылка объект range             *******************************
  function FindInDoc(const aWdDoc : OleVariant; const aSearchText : String) : oleVariant;
const
  wdFindStop = 0; //Завершить поиск при достижении границы диапазона.
var
  wdRng, wdFind : Variant;
begin
  VarClear(Result);
  //Диапазон, охватывающий всё содержимое документа.
  wdRng := aWdDoc.Content;
  //Настройка поиска.
  wdFind := wdRng.Find;
  //wdFind.ClearFormatting;
  wdFind.Text := aSearchText;
  //True - поиск вести от начала - к концу диапазона.
  wdFind.Forward := True;
  //wdFindStop - завершить поиск при достижении границы диапазона.
  //wdFind.Wrap := wdFindStop;
  //Поиск текста.
  if wdFind.Execute then Result := wdRng;
end;


//******** Процедура открытия документа EXCEL по заданному пути **************************************************************
procedure ActExcelOpenDoc;
begin

  if StartExcel then
   begin
    ShowMessage('Процесс Excel запущен');
   end;

  DIRExName:=DIRName+'xlsm';
  MyBook:= EXC.WorkBooks.Open(DIRExName);  //Открываем документ
  ShowMessage(DIRExName);   //Показывает директорию открытого файла с именем и расширением

end;


//********** Процедура запускает МАКРОС EXCEL копирующий таблицы как рисунки  ***********
//********** на листе EXCEL и вставляющий их там же, и затем процедура эти    ***********
//********** рисунки вставляет в документ WORD на заданные места путем замены ***********
//********** заданного текста                                                 ***********
procedure TableAsPicturePaste;
  var  ReplaceText : array [1..10] of string;
                 j : Integer;

  begin

       EXC.Run('ТаблКакРис');

        ShowMessage('000');
       MyWorkSheet:=MyBook.Sheets['Таблицы расч'];
        ShowMessage('001');
       MyRange:=EXC.Range['a2'];
       Book.Activate;
      // MyRange2:=EXC.Range['h1'];
     //  Book.Range(1,10).Paste;

     ReplaceText [1]:='#1';        //заменяемый на таблицы текст в WORD
     ReplaceText [2]:='#2';        //заменяемый на таблицы текст в WORD
     ReplaceText [3]:='#3';        //заменяемый на таблицы текст в WORD
     ReplaceText [4]:='#4';        //заменяемый на таблицы текст в WORD
     ReplaceText [5]:='#5';        //заменяемый на таблицы текст в WORD

     j:=1;

     repeat
          ShowMessage('002');
        Shp := MyWorkSheet.Shapes.Item(j);
        Shp.Copy;
        ShowMessage('111');
        MyRange2 := FindInDoc(Book, ReplaceText [j]);

       if VarIsClear(MyRange2) then begin
          ShowMessage('Текст НЕ найден.');
          Exit;
       end;

         MyRange2.Select;      //Выделяем найденный текст.
         MyRange2.Paste;       //Заменяем найденный текст из буфера обмена.

        j:= j+1;
     until ReplaceText [j]  = '';

   ShowMessage('рисунки из ЭКЗЕЛЬ вставлены');

  end;


procedure InsertPictureWord;
 var
 ReplaceText : array [1..5] of string;
 DirPictureDoc : string;
 i: Integer;

begin
  DirPictureDoc:= DIRFile+'доки.docx';
  BookPicture:=WRD.Documents.Open(DirPictureDoc);
  ShowMessage('файл доки открыт');
//  BookPicture.Activate;
 // wdInlineShapes := BookPicture.InlineShapes;
//  wdInlineShapes.Count;
 // ShowMessage('Количество картинок '+inttostr(wdInlineShapes.Count));
  // BookPicture.Range.InlineShapes.count;  //  error
   // BookPicture.Range.InlineShapes(BookPicture.Range.InlineShapes.count).Activate;

   ReplaceText [1]:='##1';
   ReplaceText [2]:='##2';
   ReplaceText [3]:='##3';
   ReplaceText [4]:='';

   BookPicture.Range.InlineShapes.Item(1).Range.CopyAsPicture;
   MyRange2 := FindInDoc(Book, ReplaceText [1]);
    if VarIsClear(MyRange2) then begin
        ShowMessage('Текст ##1 НЕ найден.');
        Exit;
    end;
     MyRange2.Paste;
  // MyRange2.Select;
  // Book.Activate;
  // Book.Selection.InsertAfter(' text1' );
  // MyRange2.Paste;        //OK
      MyRange2.InsertAfter('MS Word')  ;
 //  BookPicture.Range.InlineShapes.Item(2).Range.CopyAsPicture;
 //  MyRange2.Select;


     for i:=2 to 4  do
       begin
         BookPicture.Range.InlineShapes.Item(i).Range.CopyAsPicture;
         MyRange2 := FindInDoc(Book, ReplaceText [2]);
         if VarIsClear(MyRange2) then begin
            ShowMessage('Текст ##2 НЕ найден.');
            Exit;
         end;
         MyRange2.Paste;
         if i<4 then  MyRange2.InsertAfter(ReplaceText [2])  ;

       end;




        for i:=5 to 100  do
       begin
         MyRange2 := FindInDoc(Book, ReplaceText [3]);
         if VarIsClear(MyRange2) then begin
            ShowMessage('Текст ##3 НЕ найден.');
            Exit;
         end;
         try
         BookPicture.Range.InlineShapes.Item(i).Range.CopyAsPicture;



         except
         //MyRange2.Insert('')  ;
         break;
         end;
         MyRange2.Paste;
         MyRange2.InsertAfter(ReplaceText [3])  ;
       end;
      ShowMessage(IntToStr(i));
      MyRange2 := FindInDoc(Book, ReplaceText [3]);
      MyRange2.text:=ReplaceText [4];
  // MyRange2.Paste;


 {
   BookPicture.Range.InlineShapes.Item(2).Range.CopyAsPicture;
   MyRange2 := FindInDoc(Book, ReplaceText [2]);
    if VarIsClear(MyRange2) then begin
        ShowMessage('Текст 2 НЕ найден.');
        Exit;
    end;
   MyRange2.Select;
   MyRange2.Paste;

  }


  // MyRange2.Paste;
    ShowMessage('енд ');


end;


 procedure InsertObzor;
 var
    DirObzor,ReplaceTextObzor : string;

 begin
  MyWorkSheet2:=MyBook.Sheets['Ввод'];
  RangeObzor:=MyWorkSheet2.Range['b22'];
  ObzorValue:= RangeObzor.Value;
  ShowMessage(vartostr(ObzorValue));
  DirObzor:= 'Z:\GRAND NEVA\2014\ОБЗОРЫ районы\'+ObzorValue+' район.docx';
  BookObzor:=WRD.Documents.Open(DirObzor);
  BookObzor.Range.Copy;
  ReplaceTextObzor:='###1';
  MyRange2 := FindInDoc(Book, ReplaceTextObzor);
  if VarIsClear(MyRange2) then begin
            ShowMessage('Текст ###1 ReplaceTextObzor НЕ найден.');
            Exit;
         end;
  MyRange2.Paste;


 end;


 procedure InsertAnalogi;
  var
        DirAnalogDoc, ReplaceTextAnalog : string;

 begin

    DirAnalogDoc:= DIRFile+'аналоги.docx';
    BookAnalog:=WRD.Documents.Open(DirAnalogDoc);
    BookAnalog.Range.Copy;
    ReplaceTextAnalog:='####1';
    MyRange2 := FindInDoc(Book, ReplaceTextAnalog);
    if VarIsClear(MyRange2) then begin
            ShowMessage('Текст ####1 ReplaceTextAnalog НЕ найден.');
            Exit;
         end;
    MyRange2.Paste;


 end;




  procedure InsertLocation;
  var
       i: Integer;
       DirLocationDoc  : string;
       ReplaceTextLoc : array [1..3] of string;

  begin
       ReplaceTextLoc [1]:='#####1';
       ReplaceTextLoc [2]:='#####2';
       ReplaceTextLoc [3]:='#####3';


     DirLocationDoc:= DIRFile+'место.docx';
     BookLocation:=WRD.Documents.Open(DirLocationDoc);

       for i:=1 to 3  do
       begin
         BookLocation.Range.InlineShapes.Item(i).Range.CopyAsPicture;
         MyRange2 := FindInDoc(Book, ReplaceTextLoc [i]);
         if VarIsClear(MyRange2) then begin
            ShowMessage('Текст ####1 Location НЕ найден.');
            Exit;
         end;
         MyRange2.Paste;
       end;



  end;



   procedure MacroSli;

   begin
      Book.Activate;
      WRD.Run('Слияние');

   end;

   procedure CloseWordDocs;


   begin
    BookLocation.Close;
    BookAnalog.Close;
    BookObzor.Close;
    BookPicture.Close;
   end;

   procedure CloseExcel;

   begin
    EXC.DisplayAlerts := False;
    EXC.quit;
   end;




 //*****************************************************************************
 {
  procedure SaveWordDocs;

  begin

   NewWordDocDir:=DIRFName;
   Insert(ProgramName,NewWordDocDir,LengthDir+1);
   FileFormat:='wdFormatDocument';
    Book.Activate;
   Book.SaveAs(NewWordDocDir,FileFormat);

  end;
  }
  {
   procedure Proba;

   begin
    BookBlevantone:=WRD.Documents.Add;
    Book.Range.Copy;
    BookBlevantone.Range.Paste;

   end;
   }

//******************************************************************************

procedure TForm1.Button2Click(Sender: TObject);
begin

   if StartWORD then
   begin
    ShowMessage('Процесс  ВОРД запущен');
   end;

  if not OpenDialog1.Execute then Exit;  //тут открывается диалог выбора файла, и если пользователь нажал "Cancel", то выходим
 // W:=CreateOLEObject('Word.Application');  //запускается сервер автоматизации
  WRD.visible:=true;     //делаем Ворд видимым or not
 Book:= WRD.Documents.Open(OpenDialog1.FileName);
  // WRD.visible:=false;     //делаем Ворд видимым or not
  //MyRange2:=Book;

  DIRFName := OpenDialog1.FileName;
  ShowMessage(DIRFName);   //Показывает директорию открытого файла с именем и расширением
  DIRName:= DIRDetect;
  ShowMessage(DIRName);
  ShowMessage(DIRFileDetect);
  DIRFile:= DIRFileDetect;

  NewWordDocDir:=DIRFName;
  Insert(ProgramName,NewWordDocDir,LengthDir+1);
  ShowMessage(NewWordDocDir);
  //***********************************Запускаем и открываем EXCEL по заданной директории************************


   ActExcelOpenDoc;
  EXC.visible:=true;     //делаем Excel видимым
 // EXC.visible:=false;
  //EXC.Range[EXC.Cells[1, 1], EXC.Cells[5, 3]].Select;

  TableAsPicturePaste;

  InsertPictureWord;
  InsertObzor ;
  InsertAnalogi;
  InsertLocation;
  MacroSli;


  CloseWordDocs;
  CloseExcel;

  WRD.visible:=true;     //делаем Ворд видимым
  ShowMessage('program end');
  //Book.Range(1,10).Paste;



end;

 {
procedure DIRDetect;

begin

end;
   }



begin
   ProgramName:='Blevantone';





end.


