unit WordActivate;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.StdCtrls, System.Win.ComObj, Winapi.ActiveX, TlHelp32, ShellAPI,
  Vcl.Samples.Spin, Vcl.ComCtrls, Vcl.Menus, Vcl.ExtCtrls, System.IniFiles;


type
  TForm1 = class(TForm)
    Button2: TButton;
    OpenDialog1: TOpenDialog;
    Label1: TLabel;
    ProgressBar1: TProgressBar;
    Label2: TLabel;
    MainMenu1: TMainMenu;
    MenuFile: TMenuItem;
    MenuHelp: TMenuItem;
    AboutProgram: TMenuItem;
    WhatDo: TMenuItem;
    Author: TMenuItem;
    ExitP: TMenuItem;
    Help1: TMenuItem;
    Help2: TMenuItem;
    Help3: TMenuItem;
    Start: TMenuItem;
    Image1: TImage;
    Button1: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    Button7: TButton;

    procedure Button2Click(Sender: TObject);
    procedure AuthorClick(Sender: TObject);
    procedure WhatDoClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);


  //  procedure Label1Click(Sender: TObject);
  //  procedure Label2Click(Sender: TObject);




  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

  WRD,Book,BookPicture,BookAnalogiCoInvest,BookZad,RangeBookZad,BookObzor,BookObzorRec,BookObzorDoci,BookObzorRF,
      BookObzorRegion, BookObzorObj, BookObzorFco, BookObzorOgr, BookObzorDop, BookObzorTerm, BookObzorLit,
      BookAnalog,BookAnalogBuild,BookLocation,BookFoto,
      BookBlevantone,RangeBookBlevantone,wdInlineShapes,wdInlineShapes2: OleVariant;
  EXC,MyBook,MyWorkSheet,MyWorkSheet2,MyRange,MyRange2,RangeObzor,RangeObzorObj,
      Shp,ShpWrd,ShpWrd2,ShpWrd3,vstart,vend: OleVariant;
  var W,ObzorValue,ObzorValueObj,ObzorValueRF, ObzorValueDoci, ObzorValueLit,ObzorValueRec, ObzorValueRegion,
      ObzorValueFco, ObzorValueOgr, ObzorValueDop, ObzorValueTerm :variant;
                 i:Integer;
        LengthDir : Integer;
        ProgBar   : Integer;
  var DIRFName    : string;
  var DIRName     : string;
  var  DirExName  : string;
  var NewWordDocDir: string;
  var FileFormat  : OleVariant;
  var ProgramName : string;
  var  DIRFile    : string;
 // var DirFolderDetect : string;
  var  Label2C     : string;
  var  IniFile: TIniFile;
  var  IniSectionValue : array [1..20,1..3] of string;
  var Password    : string;
 // var Dialog : OpenDialog1;

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
 i ,L       : Integer;
 DIRECTORY: string;

 begin


        DIRECTORY:=DIRFname;
   {
        while (Smb<>'.') do
          begin
          Smb:=DIRFname[i];
          DIRECTORY:=DIRECTORY+Smb;
          i:=i+1;
          end;
      }

       L:=Length(DIRFname);
        i:=L;
        Smb:=DIRFname[i];

        while (Smb<>'.') do
          begin
             i:=i-1;
             Smb:=DIRFname[i];
          end;

          Delete(DIRECTORY, i+1, l-i);


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

 //****************** определяется директория папки без названия файла  **************
 //******************               со входящим параметром              **************
 function DIRFileDetect2(DIRFname2 : string): string;

 var

 Smb         : Char;
 i , L       : Integer;
 DIRECTORY   : string;


 begin
        DIRECTORY:=' ';
        DIRECTORY:=DIRFname2;
        L:=Length(DIRFname2);
        i:=L;
        Smb:=DIRFname2[i];

        while (Smb<>'\') do
          begin
             i:=i-1;
             Smb:=DIRFname2[i];
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
  wdRng, wdFind : OleVariant;
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


//******** Процедура открытия документа EXCEL по заданному пути **********************

procedure ActExcelOpenDoc;
begin

  DIRExName:=DIRName+'xlsm';

  MyBook:= EXC.WorkBooks.Open(DIRExName);  //Открываем документ
//  ShowMessage(DIRExName);   //Показывает директорию открытого файла с именем и расширением

end;



//******** Процедура определения пути к документу EXCEL    *****************************
//********         по запросу Экспертного Заключения       ******************************

procedure DetectExcelOpenExpertDoc;

 var
 CountDir  : Integer;

begin

 // ShowMessage(DIRFileDetect);

  CountDir := Length(DIRFileDetect);

  Delete(DIRName, CountDir+1, 2);

  DIRExName:=DIRName+'xlsm';

 // ShowMessage(DIRExName);   //Показывает директорию открытого файла с именем и расширением

end;


//**********              ДЛЯ КВАРТИР, КОМНАТ, ЗУ, ТАУНХАУЗОВ                  ***********
//********** Процедура запускает МАКРОС EXCEL копирующий таблицы как рисунки  ***********
//********** на листе EXCEL и вставляющий их там же, и затем процедура эти    ***********
//********** рисунки вставляет в документ WORD на заданные места путем замены ***********
//********** заданного текста                                                 ***********
procedure TableAsPicturePaste(Book : OleVariant);
  var  ReplaceText : array [1..10] of string;
                 j : Integer;

  begin
   //   ShowMessage('0');
    //  EXC.Run('PERSONAL.XLSB!ТаблКакРис');
    //  EXC.Run('ТКР');
    //  EXC.Run('ТаблКакРис');
   //  ShowMessage('000');
   //  Sleep(1000);
     MyWorkSheet:=MyBook.Sheets['Таблицы расч'];

   //  ShowMessage('001');
   //  Sleep(500);
     Shp := MyWorkSheet.Shapes.Item(1);
    //  MyRange:=EXC.Range['a2'];
    //   Book.Activate;
    // MyRange2:=EXC.Range['h1'];
    //  Book.Range(1,10).Paste;
    //  ShowMessage('0002');
     ReplaceText [1]:='1#1';        //заменяемый на таблицы текст в WORD
     ReplaceText [2]:='1#2';        //заменяемый на таблицы текст в WORD
     ReplaceText [3]:='1#3';        //заменяемый на таблицы текст в WORD
     ReplaceText [4]:='1#4';        //заменяемый на таблицы текст в WORD
     ReplaceText [5]:='1#5';        //заменяемый на таблицы текст в WORD
     ReplaceText [6]:='1#6';        //заменяемый на таблицы текст в WORD

     j:=1;

     repeat
      //ShowMessage('002');
        Shp := MyWorkSheet.Shapes.Item(j);
        Shp.Copy;
      //ShowMessage('111');
        MyRange2 := FindInDoc(Book, ReplaceText [j]);

       if VarIsClear(MyRange2) then begin
          ShowMessage('Текст НЕ найден.');
          Exit;
       end;

         MyRange2.Select;      //Выделяем найденный текст.
         MyRange2.Paste;       //Заменяем найденный текст из буфера обмена.

        j:= j+1;
     until ReplaceText [j]  = '';

   //ShowMessage('рисунки из ЭКЗЕЛЬ вставлены');

  end;


//**********              ДЛЯ ЗДАНИЙ С ЗЕМЕЛЬНЫМ УЧАСТКОМ                     ***********
//********** Процедура запускает МАКРОС EXCEL копирующий таблицы как рисунки  ***********
//********** на листе EXCEL и вставляющий их там же, и затем процедура эти    ***********
//********** рисунки вставляет в документ WORD на заданные места путем замены ***********
//********** заданного текста                                                 ***********

procedure TableAsPicturePasteBuildLand(Book : OleVariant);
  var  ReplaceText : array [1..30] of string;
                 j : Integer;

  begin
   //   ShowMessage('0');
    //  EXC.Run('PERSONAL.XLSB!ТаблКакРис');
    //  EXC.Run('ТКР');
    //  EXC.Run('ТаблКакРис');
   //  ShowMessage('000');
   //  Sleep(1000);
     MyWorkSheet:=MyBook.Sheets['Таблицы расч'];

   //  ShowMessage('001');
   //  Sleep(500);
     Shp := MyWorkSheet.Shapes.Item(1);
    //  MyRange:=EXC.Range['a2'];
    //   Book.Activate;
    // MyRange2:=EXC.Range['h1'];
    //  Book.Range(1,10).Paste;
    //  ShowMessage('0002');
     ReplaceText [1]:='1#1ЗУ';        //заменяемый на таблицы текст в WORD
     ReplaceText [2]:='1#2ЗУ';        //заменяемый на таблицы текст в WORD
     ReplaceText [3]:='1#3ЗУ';        //заменяемый на таблицы текст в WORD
     ReplaceText [4]:='1#4ЗУ';        //заменяемый на таблицы текст в WORD
     ReplaceText [5]:='1#5ЗУ';        //заменяемый на таблицы текст в WORD
     ReplaceText [6]:='1#6ЗУ';        //заменяемый на таблицы текст в WORD

     ReplaceText [7]:='1#1ЗДЗУ';        //заменяемый на таблицы текст в WORD
     ReplaceText [8]:='1#2ЗДЗУ';        //заменяемый на таблицы текст в WORD
     ReplaceText [9]:='1#3ЗДЗУ';        //заменяемый на таблицы текст в WORD
     ReplaceText [10]:='1#4ЗДЗУ';        //заменяемый на таблицы текст в WORD
     ReplaceText [11]:='1#5ЗДЗУ';        //заменяемый на таблицы текст в WORD
     ReplaceText [12]:='1#6ЗДЗУ';        //заменяемый на таблицы текст в WORD

  //   ReplaceText [13]:='1#1ЗАТР';      //заменяемый на таблицы текст в WORD
  //   ReplaceText [14]:='1#2ЗАТР';      //заменяемый на таблицы текст в WORD
     ReplaceText [13]:='1#3ЗАТР';        //заменяемый на таблицы текст в WORD
     ReplaceText [14]:='1#4ЗАТР';        //заменяемый на таблицы текст в WORD
     ReplaceText [15]:='1#5ЗАТР';        //заменяемый на таблицы текст в WORD
     ReplaceText [16]:='1#6ЗАТР';        //заменяемый на таблицы текст в WORD
     ReplaceText [17]:='1#7ЗАТР';        //заменяемый на таблицы текст в WORD
     ReplaceText [18]:='1#8ЗАТР';        //заменяемый на таблицы текст в WORD
     ReplaceText [19]:='1#9ЗАТР';        //заменяемый на таблицы текст в WORD
     ReplaceText [20]:='1#10ЗАТР';       //заменяемый на таблицы текст в WORD

     ReplaceText [21]:='1#1СОГЛ';        //заменяемый на таблицы текст в WORD
     ReplaceText [22]:='1#2СОГЛ';        //заменяемый на таблицы текст в WORD
     ReplaceText [23]:='1#3СОГЛ';        //заменяемый на таблицы текст в WORD
     ReplaceText [24]:='1#4СОГЛ';        //заменяемый на таблицы текст в WORD

     j:=1;

     repeat
      //ShowMessage('002');
        Shp := MyWorkSheet.Shapes.Item(j);
        Shp.Copy;
      //ShowMessage('111');
        MyRange2 := FindInDoc(Book, ReplaceText [j]);

       if VarIsClear(MyRange2) then begin
          ShowMessage('Текст НЕ найден.');
          Exit;
       end;

         MyRange2.Select;      //Выделяем найденный текст.
         MyRange2.Paste;       //Заменяем найденный текст из буфера обмена.

        j:= j+1;
     until ReplaceText [j]  = '';

   //ShowMessage('рисунки из ЭКЗЕЛЬ вставлены');

  end;


 //***************ВСТАВКА РИСУНКОВ АКТ ОСМОТРА И ЗАДАНИЕ ИЗ EXCEL В ЗАДАКТ*********************************************************************


 //******* Процедура вставляет  рисунки акт осмотра и задание из документа EXCEL   ********
//******* в документ в WORD   под названием задакт            ********
//******* (          ********

 procedure InsertPictureWordZad;
  var
    j,i : Integer;
 begin

     BookZad:=WRD.Documents.Open(DIRFile+'задакт.docx');

     MyWorkSheet:=MyBook.Sheets['рисунки'];
     j:=MyWorkSheet.Shapes.Count;
  //   ShowMessage(IntToStr(j));

     RangeBookZad:=BookZad.Range;
     for i := 1 to j do
     begin
       Shp := MyWorkSheet.Shapes.Item(i);
       Shp.Copy;

       RangeBookZad.InsertAfter('Error 404');
       RangeBookZad := FindInDoc(BookZad, 'Error 404');

       RangeBookZad.Paste;
     end;


 end;




//************ВСТАВКА ИЗ ДОКИ В ОТЧЕТ WORD*****************************************


//******* Процедура вставляет  рисунки из документа WORD   ********
//******* в другой документ в заданные места               ********
//******* (в данном случае это Зад. и Акт осмотра          ********
//procedure InsertPictureWord(Book : OleVariant);
 procedure InsertPictureWord();
 var
 ReplaceText   : array [1..5] of string;
 ReplaceTextValue : string;
 DirPictureDoc : string;
 i,j             : Integer;

begin
  DirPictureDoc:= DIRFile+'доки.docx';
  BookPicture:=WRD.Documents.Open(DirPictureDoc);


   ReplaceText [1]:='1##1';
   ReplaceText [2]:='1##2';
   ReplaceText [3]:='1##3';
   ReplaceText [4]:='';
   j:=0;

    j:=BookPicture.InlineShapes.Count;

     case j of
      0 : ShowMessage('В доках нет сканов');
      1 : ShowMessage('В доках нет акта осмотра и остальных документов');
      2 : ShowMessage('В доках не хватает сканов документов и акта осмотра');
      3 : ShowMessage('В доках не хватает сканов документов и акта осмотра');
      4 : ShowMessage('В доках нет сканов документов');
     end;
   if j=0  then   begin
                      BookPicture.Close;
                      Exit;
   end;






    for i:=1 to j  do
       begin

             ShpWrd2:= BookPicture.Range.InlineShapes;
             ShpWrd2.Item(i).select;
             WRD.selection.Copyaspicture;

             ReplaceTextValue := ReplaceText[3];
             case i of
             1 :  ReplaceTextValue := ReplaceText[1];
             2 :  ReplaceTextValue := ReplaceText[2];
             3 :  ReplaceTextValue := ReplaceText[2];
             4 :  ReplaceTextValue := ReplaceText[2];
             5 :  ReplaceTextValue := ReplaceText[3];
             end;


             MyRange2 := FindInDoc(Book, ReplaceTextValue);

         if  VarIsClear(MyRange2) then begin
             ShowMessage('Текст в документе;'+ReplaceTextValue+' НЕ найден.');
             BookPicture.Close;
             Exit;
         end;


            MyRange2.Paste;

 //*************************************
         if i=1 then
            begin
           ProgBar:=43;
           Form1.ProgressBar1.Position := ProgBar ;
           Label2C:='вставил Задание, пихаю Акт осмотра в документ';
           Form1.Label2.Caption:=Label2C;
          end;

         if i=4 then
           begin
           ProgBar:=50;
           Form1.ProgressBar1.Position := ProgBar ;
           Label2C:='вставил Акт, пихаю остальную лабуду из доки';
           Form1.Label2.Caption:=Label2C;
          end;

         if i=12 then
            begin
             ProgBar:=54;
             Form1.ProgressBar1.Position := ProgBar ;
             Label2C:='не волнуйся, работаю, пихаю...';
             Form1.Label2.Caption:=Label2C;
            end;

 //**********************************************

         if (i<j) xor (i<>1) xor (i<>4) then  MyRange2.InsertAfter(ReplaceTextValue)  ;

       end;

       BookPicture.Close


end;



//********************ВСТАВЛЯЕТ РАЙОН ПО НАЗВАНИЮ В EXCEL В ОТЧЕТ WORD**************************************************************************

//**** Процедура ищет из таблицы EXCEL название района               *****
//**** открывает документ WORD с обзором по заданной директрии       *****
//**** и вставляет в документ WORD отчета                            *****

 procedure InsertObzor();
 var
    DirObzor,ReplaceTextObzor : string;
    DefaultReadIniFile        : string;

 begin
      MyWorkSheet2:=MyBook.Sheets['Ввод'];
      RangeObzor:=MyWorkSheet2.Range['b22'];
      ObzorValue:= RangeObzor.Value;

    //ShowMessage(vartostr(ObzorValue));

        DefaultReadIniFile := 'Z:\GRAND NEVA\2014\ОБЗОРЫ районыZ\';   //Значение директории обзора по умолчанию
        DirObzor:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [1,2],'Заглушка по факту стринговая')+ObzorValue+' район.docx';

     //ShowMessage(DirObzor);

      try
      BookObzor:=WRD.Documents.Open(DirObzor);
      except

         ShowMessage('BLEVANTONE не смог открыть файл с обзорами.'+
       ' Пожалуйста укажите расположение документа с обзором района '+
        ObzorValue+' ну или, если нет, любого другого ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //тут открывается диалог выбора файла, и если пользователь нажал "Cancel", то выходим

          DirObzor := Form1.OpenDialog1.FileName;
          BookObzor:=WRD.Documents.Open(DirObzor);

         //******* Определение и запись в INI файл директории папки с обзорами *****

         IniSectionValue [1,3]:= DIRFileDetect2(DirObzor);  //отсечение имени файла с обзором района от пути
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [1,2], IniSectionValue [1,3]);

      end;

      BookObzor.Range.Copy;
      ReplaceTextObzor:='1###1';
      MyRange2 := FindInDoc(Book, ReplaceTextObzor);

         if VarIsClear(MyRange2) then begin
            ShowMessage('Текст 1###1 ReplaceTextObzor НЕ найден.');
            Exit;
         end;

  MyRange2.Paste;
 end;



 //*********************ВСТАВЛЯЕТ ОБЗОР ОБЬЕКТА ПО НАЗВАНИЮ В EXCEL В ОТЧЕТ WORD************************************

//**** Процедура ищет из таблицы EXCEL название объекта оценки  obj             *****
//**** открывает документ WORD с обзором объектов по заданной директрии       *****
//**** и вставляет в документ WORD отчета                            *****
 procedure InsertObzorObj();
 var
    DirObzorObj,ReplaceTextObzorObj : string;
    DefaultReadIniFileObj        : string;

 begin
      MyWorkSheet2:=MyBook.Sheets['Ввод'];
      RangeObzorObj:=MyWorkSheet2.Range['c3'];
      ObzorValueObj:= RangeObzorObj.Value;

    //ShowMessage(vartostr(ObzorValue));

        DefaultReadIniFileObj := 'Z:\GRAND NEVA\2014\ОБЗОРЫ районыZ\';   //Значение директории обзора по умолчанию
        DirObzorObj:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [4,2], 'стринговая заглушка')+ObzorValueObj+'.docx';

     //ShowMessage(DirObzorObj);

      try
      BookObzorObj:=WRD.Documents.Open(DirObzorObj);
      except

         ShowMessage('BLEVANTONE не смог открыть файл с обзором объекта.'+
       ' Пожалуйста укажите расположение документа с обзором обьекта '+
        ObzorValueObj+' ну или, если нет, любого другого ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //тут открывается диалог выбора файла, и если пользователь нажал "Cancel", то выходим

          DirObzorObj := Form1.OpenDialog1.FileName;
          BookObzorObj:=WRD.Documents.Open(DirObzorObj);

         //******* Определение и запись в INI файл директории папки с обзорами *****

         IniSectionValue [4,3]:= DIRFileDetect2(DirObzorObj);  //отсечение имени файла с обзором района от пути
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [4,2], IniSectionValue [4,3]);

      end;

      BookObzorObj.Range.Copy;
      ReplaceTextObzorObj:='7###obj';
      MyRange2 := FindInDoc(Book, ReplaceTextObzorObj);

         if VarIsClear(MyRange2) then begin
            ShowMessage('Текст 7###obj ReplaceTextObzorObj НЕ найден.');
            Exit;
         end;

  MyRange2.Paste;
 end;





//********************ВСТАВЛЯЕТ ОБЗОР РФ И РЕГИОН В ОТЧЕТ WORD******************************************************


//**** Процедура ищет из таблицы EXCEL название объекта оценки  obj             *****
//**** открывает документ WORD с обзором объектов по заданной директрии       *****
//**** и вставляет в документ WORD отчета                            *****
 procedure InsertObzorRFandRegion();
 var
    DirObzorRF, DirObzorRegion, ReplaceTextObzorRF : string;
    ReplaceTextObzorRegion : string;
    DefaultReadIniFileRF,DefaultReadIniFileRegion         : string;

 begin
     // MyWorkSheet2:=MyBook.Sheets['Ввод'];
     // RangeObzorObj:=MyWorkSheet2.Range['c3'];
      ObzorValueRF:= 'обзор рф';
      ObzorValueRegion:= 'обзор регион';

    //ShowMessage(vartostr(ObzorValue));

        DefaultReadIniFileRF := 'Z:\GRAND NEVA\2014\ОБЗОРЫ районыZ\';   //Значение директории обзора по умолчанию
        DirObzorRF:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [5,2], 'стринговая заглушка')+ObzorValueRF+'.docx';


        DefaultReadIniFileRegion := 'Z:\GRAND NEVA\2014\ОБЗОРЫ районыZ\';   //Значение директории обзора по умолчанию
        DirObzorRegion:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [6,2], 'стринговая заглушка')+ObzorValueRegion+'.docx';

     //ShowMessage(DirObzorObj);

      try
      BookObzorRF:=WRD.Documents.Open(DirObzorRF);
      except

         ShowMessage('BLEVANTONE не смог открыть файл с обзором RF.'+
       ' Пожалуйста укажите расположение документа с обзором RF '+
        ObzorValueRF+' ну или, если нет, любого другого ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //тут открывается диалог выбора файла, и если пользователь нажал "Cancel", то выходим

          DirObzorRF := Form1.OpenDialog1.FileName;
          BookObzorRF:=WRD.Documents.Open(DirObzorRF);

         //******* Определение и запись в INI файл директории папки с обзорами *****

         IniSectionValue [5,3]:= DIRFileDetect2(DirObzorRF);  //отсечение имени файла с обзором района от пути
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [5,2], IniSectionValue [5,3]);

      end;



      try
      BookObzorRegion:=WRD.Documents.Open(DirObzorRegion);
      except

         ShowMessage('BLEVANTONE не смог открыть файл с обзором Region.'+
       ' Пожалуйста укажите расположение документа с обзором Region '+
        ObzorValueRegion+' ну или, если нет, любого другого ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //тут открывается диалог выбора файла, и если пользователь нажал "Cancel", то выходим

          DirObzorRegion := Form1.OpenDialog1.FileName;
          BookObzorRegion:=WRD.Documents.Open(DirObzorRegion);

         //******* Определение и запись в INI файл директории папки с обзорами *****

         IniSectionValue [6,3]:= DIRFileDetect2(DirObzorRegion);  //отсечение имени файла с обзором района от пути
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [6,2], IniSectionValue [6,3]);

      end;


      BookObzorRF.Range.Copy;
      ReplaceTextObzorRF:='8###rf';
      MyRange2 := FindInDoc(Book, ReplaceTextObzorRF);

         if VarIsClear(MyRange2) then begin
            ShowMessage('Текст 8###rf ReplaceTextObzorRF НЕ найден.');
            Exit;
         end;

  MyRange2.Paste;


  BookObzorRegion.Range.Copy;
      ReplaceTextObzorRegion:='9###reg';
      MyRange2 := FindInDoc(Book, ReplaceTextObzorRegion);

         if VarIsClear(MyRange2) then begin
            ShowMessage('Текст 9###reg ReplaceTextObzorRegion НЕ найден.');
            Exit;
         end;

  MyRange2.Paste;
 end;



//*****************ДОКУМЕНТЫ ОЦЕНОЧНЫЕ СКАНЫ ВСТАВЛЯЕТ В ОТЧЕТ WORD******************************************


//**** Процедура ищет из файла ini путь к докам ЧПО или ООО (страховка, квалаттестат и др. *****
//**** открывает документ WORD с документами (страх, квал и др) по заданной директрии       *****
//**** и вставляет в документ WORD отчета                            *****
 procedure InsertDociOcen();
 var
    DirDoci,ReplaceTextDoci : string;
    DefaultReadIniFileDoci        : string;

 begin
     // MyWorkSheet2:=MyBook.Sheets['Ввод'];
     // RangeObzorObj:=MyWorkSheet2.Range['c3'];
      ObzorValueDoci:= 'ЧПО Пичукан';

    //ShowMessage(vartostr(ObzorValue));

        DefaultReadIniFileDoci := 'Z:\GRAND NEVA\2014\ОБЗОРЫ районыZ\';   //Значение директории обзора по умолчанию
        DirDoci:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [7,2], 'стринговая заглушка');

     //ShowMessage(DirObzorObj);

      try
      BookObzorDoci:=WRD.Documents.Open(DirDoci);
      except

         ShowMessage('BLEVANTONE не смог открыть файл с документами.'+
       ' Пожалуйста укажите расположение документа с документами '+
        ObzorValueDoci+' ну или, если нет, любого другого ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //тут открывается диалог выбора файла, и если пользователь нажал "Cancel", то выходим

          DirDoci := Form1.OpenDialog1.FileName;
          BookObzorDoci:=WRD.Documents.Open(DirDoci);

         //******* Определение и запись в INI файл директории папки с обзорами *****

         IniSectionValue [7,3]:= DirDoci;  //нет отсечения имени файла с обзором района от пути
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [7,2], IniSectionValue [7,3]);

      end;

      BookObzorDoci.Range.Copy;
      ReplaceTextDoci:='10###doc';
      MyRange2 := FindInDoc(Book, ReplaceTextDoci);

         if VarIsClear(MyRange2) then begin
            ShowMessage('Текст 10###doc ReplaceTextDoci НЕ найден.');
            Exit;
         end;

  MyRange2.Paste;
 end;






 //****************ВСТАВЛЯЕТ В ОТЧЕТ РЕКВИЗИТЫ ОЦЕНОЧНЫЕ**********************************************************

//**** Процедура ищет из файла ini путь к докам ЧПО или ООО (страховка, квалаттестат и др. *****
//**** открывает документ WORD с документами (страх, квал и др) по заданной директрии       *****
//**** и вставляет в документ WORD отчета                            *****
 procedure InsertDociRec();
 var
    DirRec,ReplaceTextRec : string;
    DefaultReadIniFileRec        : string;

 begin
     // MyWorkSheet2:=MyBook.Sheets['Ввод'];
     // RangeObzorObj:=MyWorkSheet2.Range['c3'];
      ObzorValueRec:= 'ЧПО Пичукан реквизиты';

    //ShowMessage(vartostr(ObzorValue));

        DefaultReadIniFileRec := 'Z:\GRAND NEVA\2014\ОБЗОРЫ районыZ\';   //Значение директории обзора по умолчанию
        DirRec:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [8,2], 'стринговая заглушка');

     //ShowMessage(DirObzorObj);

      try
      BookObzorRec:=WRD.Documents.Open(DirRec);
      except

         ShowMessage('BLEVANTONE не смог открыть файл с реквизитами оценщика.'+
       ' Пожалуйста укажите расположение документа с реквизитами '+
        ObzorValueRec+' ну или, если нет, любого другого ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //тут открывается диалог выбора файла, и если пользователь нажал "Cancel", то выходим

          DirRec := Form1.OpenDialog1.FileName;
          BookObzorRec:=WRD.Documents.Open(DirRec);

         //******* Определение и запись в INI файл директории папки с обзорами *****

         IniSectionValue [8,3]:= DirRec;  //нет отсечения имени файла с обзором района от пути
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [8,2], IniSectionValue [8,3]);

      end;

      BookObzorRec.Range.Copy;
      ReplaceTextRec:='11###rec';
      MyRange2 := FindInDoc(Book, ReplaceTextRec);

         if VarIsClear(MyRange2) then begin
            ShowMessage('Текст 11###rec в ReplaceTextRec НЕ найден.');
            Exit;
         end;

  MyRange2.Paste;
 end;




 //****************ВСТАВЛЯЕТ ФСО И РЕГЛАМЕНТНЫЕ В ОТЧЕТ WORD********************************************************************************************************


//**** Процедура ищет из файла ini путь к документу с Федеральными стандартами оценки (законодательство) *****
//**** открывает документ WORD с законодательством по заданной директрии       *****
//**** и вставляет в документ WORD отчета                            *****
 procedure InsertFco();
 var
    DirFco,ReplaceTextFco : string;
    DefaultReadIniFileFco        : string;

 begin
     // MyWorkSheet2:=MyBook.Sheets['Ввод'];
     // RangeObzorObj:=MyWorkSheet2.Range['c3'];
      ObzorValueFco:= 'ФСО';

    //ShowMessage(vartostr(ObzorValue));

        DefaultReadIniFileFco := 'Z:\GRAND NEVA\2014\ОБЗОРЫ районыZ\';   //Значение директории обзора по умолчанию
        DirFco:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [9,2], 'стринговая заглушка');

     //ShowMessage(DirObzorObj);

      try
      BookObzorFco:=WRD.Documents.Open(DirFco);
      except

         ShowMessage('BLEVANTONE не смог открыть файл с ФСО.'+
       ' Пожалуйста укажите расположение документа с ФСО '+
        ObzorValueFco+' ну или, если нет, любого другого ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //тут открывается диалог выбора файла, и если пользователь нажал "Cancel", то выходим

          DirFco := Form1.OpenDialog1.FileName;
          BookObzorFco:=WRD.Documents.Open(DirFco);

         //******* Определение и запись в INI файл директории папки с обзорами *****

         IniSectionValue [9,3]:= DirFco;  //нет отсечения имени файла с обзором района от пути
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [9,2], IniSectionValue [9,3]);

      end;

      BookObzorFco.Range.Copy;

     // ShowMessage('ФСО в буфере обменга');

      ReplaceTextFco:='ФСО###';
      MyRange2 := FindInDoc(Book, ReplaceTextFco);

         if VarIsClear(MyRange2) then begin
            ShowMessage('Текст ФСО### в ReplaceTextFco НЕ найден.');
            Exit;
         end;

  MyRange2.Paste;

  //ShowMessage('MyRange2 ФСО значение');

 end;





//********************ВСТАВЛЯЕТ СПИСОК ЛИТЕРАТУРЫ В ОТЧЕТ WORD*********************************************************************************************


//**** Процедура ищет из файла ini путь к документу со списком используемой литературы *****
//**** открывает документ WORD с списком литературы по заданной директрии              *****
//**** и вставляет в документ WORD отчета на заданное место по метке                   *****
 procedure InsertLit();
 var
    DirLit,ReplaceTextLit : string;
    DefaultReadIniFileLit        : string;

 begin
     // MyWorkSheet2:=MyBook.Sheets['Ввод'];
     // RangeObzorObj:=MyWorkSheet2.Range['c3'];
      ObzorValueLit:= 'литература';

    //ShowMessage(vartostr(ObzorValue));

        DefaultReadIniFileLit := 'Z:\GRAND NEVA\2014\ЛИТЕРАТУРА\';   //Значение директории обзора по умолчанию
        DirLit:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [13,2], 'стринговая заглушка');

     //ShowMessage(DirObzorObj);

      try
      BookObzorLit:=WRD.Documents.Open(DirLit);
      except

         ShowMessage('BLEVANTONE не смог открыть файл с литературой.'+
       ' Пожалуйста укажите расположение документа с литературой '+
        ObzorValueLit+' ну или, если нет, любого другого ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //тут открывается диалог выбора файла, и если пользователь нажал "Cancel", то выходим

          DirLit := Form1.OpenDialog1.FileName;
          BookObzorLit:=WRD.Documents.Open(DirLit);

         //******* Определение и запись в INI файл директории папки с обзорами *****

         IniSectionValue [13,3]:= DirLit;  //нет отсечения имени файла с обзором района от пути
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [13,2], IniSectionValue [13,3]);

      end;

      BookObzorLit.Range.Copy;

     // ShowMessage('литература в буфере обменга');

      ReplaceTextLit:='литер###';
      MyRange2 := FindInDoc(Book, ReplaceTextLit);

         if VarIsClear(MyRange2) then begin
            ShowMessage('Текст литер### в ReplaceTextLit НЕ найден.');
            Exit;
         end;

  MyRange2.Paste;

  //ShowMessage('MyRange2 lit значение');

 end;




//********************ВСТАВЛЯЕТ ОГРАНИЧЕНИЯ И ПРЕДПОЛОЖЕНИЯ В ОТЧЕТ WORD*******************************************************************************************


//**** Процедура ищет из файла ini путь к тексту по ограничениям и преположениям  *****
//**** открывает документ WORD с ограничениями и предположениями по заданной директрии       *****
//**** и вставляет в документ WORD отчета                            *****
 procedure InsertOgr();
 var
    DirOgr,ReplaceTextOgr : string;
    DefaultReadIniFileOgr        : string;

 begin
     // MyWorkSheet2:=MyBook.Sheets['Ввод'];
     // RangeObzorObj:=MyWorkSheet2.Range['c3'];
      ObzorValueOgr:= 'ограничения и предположения';

    //ShowMessage(vartostr(ObzorValue));

        DefaultReadIniFileOgr := 'Z:\GRAND NEVA\2014\ОБЗОРЫ районыZ\';   //Значение директории обзора по умолчанию
        DirOgr:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [10,2], 'стринговая заглушка');

     //ShowMessage(DirObzorObj);

      try
      BookObzorOgr:=WRD.Documents.Open(DirOgr);
      except

         ShowMessage('BLEVANTONE не смог открыть файл с ограничениями.'+
       ' Пожалуйста укажите расположение документа с ограничениями '+
        ObzorValueOgr+' ну или, если нет, любого другого ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //тут открывается диалог выбора файла, и если пользователь нажал "Cancel", то выходим

          DirOgr := Form1.OpenDialog1.FileName;
          BookObzorOgr:=WRD.Documents.Open(DirOgr);

         //******* Определение и запись в INI файл директории папки с обзорами *****

         IniSectionValue [10,3]:= DirOgr;  //нет отсечения имени файла с обзором района от пути
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [10,2], IniSectionValue [10,3]);

      end;

      BookObzorOgr.Range.Copy;
      ReplaceTextOgr:='ограничения###';
      MyRange2 := FindInDoc(Book, ReplaceTextOgr);

         if VarIsClear(MyRange2) then begin
            ShowMessage('Текст ограничения### в ReplaceTextOgr НЕ найден.');
            Exit;
         end;

  MyRange2.Paste;
 end;




//***************ВСТАВЛЯЕТ ДОПУЩЕНИЯ В ОТЧЕТ WORD***********************************************************


//**** Процедура ищет из файла ini путь к тексту по допущениям   *****
//**** открывает документ WORD с допущениями по заданной директрии       *****
//**** и вставляет в документ WORD отчета                            *****
 procedure InsertDop();
 var
    DirDop,ReplaceTextDop : string;
    DefaultReadIniFileDop        : string;

 begin
     // MyWorkSheet2:=MyBook.Sheets['Ввод'];
     // RangeObzorObj:=MyWorkSheet2.Range['c3'];
      ObzorValueDop:= 'особые допущения';

    //ShowMessage(vartostr(ObzorValue));

        DefaultReadIniFileDop := 'Z:\GRAND NEVA\2014\ОБЗОРЫ районыZ\';   //Значение директории обзора по умолчанию
        DirDop:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [11,2], 'стринговая заглушка');

     //ShowMessage(DirObzorObj);

      try
      BookObzorDop:=WRD.Documents.Open(DirDop);
      except

         ShowMessage('BLEVANTONE не смог открыть файл с допущениями.'+
       ' Пожалуйста укажите расположение документа с допущениями '+
        ObzorValueDop+' ну или, если нет, любого другого ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //тут открывается диалог выбора файла, и если пользователь нажал "Cancel", то выходим

          DirDop := Form1.OpenDialog1.FileName;
          BookObzorDop:=WRD.Documents.Open(DirDop);

         //******* Определение и запись в INI файл директории папки с обзорами *****

         IniSectionValue [11,3]:= DirDop;  //нет отсечения имени файла с обзором района от пути
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [11,2], IniSectionValue [11,3]);

      end;

      BookObzorDop.Range.Copy;
      ReplaceTextDop:='допущения###';
      MyRange2 := FindInDoc(Book, ReplaceTextDop);

         if VarIsClear(MyRange2) then begin
            ShowMessage('Текст допущения### в ReplaceTextDop НЕ найден.');
            Exit;
         end;

  MyRange2.Paste;
 end;





//**************ВСТАВЛЯЕТ ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ В ОТЧЕТ WORD*******************************************


//**** Процедура ищет из файла ini путь к тексту по термины и определения   *****
//**** открывает документ WORD с терминами по заданной директрии       *****
//**** и вставляет в документ WORD отчета                            *****
 procedure InsertTerm();
 var
    DirTerm,ReplaceTextTerm : string;
    DefaultReadIniFileTerm        : string;

 begin
     // MyWorkSheet2:=MyBook.Sheets['Ввод'];
     // RangeObzorObj:=MyWorkSheet2.Range['c3'];
      ObzorValueTerm:= 'термины и определения';

    //ShowMessage(vartostr(ObzorValue));

        DefaultReadIniFileTerm := 'Z:\GRAND NEVA\2014\ОБЗОРЫ районыZ\';   //Значение директории обзора по умолчанию
        DirTerm:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [12,2], 'стринговая заглушка');

     //ShowMessage(DirObzorObj);

      try
      BookObzorTerm:=WRD.Documents.Open(DirTerm);
      except

         ShowMessage('BLEVANTONE не смог открыть файл с терминами и определениями.'+
       ' Пожалуйста укажите расположение документа с терминами и определениями '+
        ObzorValueTerm+' ну или, если нет, любого другого ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //тут открывается диалог выбора файла, и если пользователь нажал "Cancel", то выходим

          DirTerm := Form1.OpenDialog1.FileName;
          BookObzorTerm:=WRD.Documents.Open(DirTerm);

         //******* Определение и запись в INI файл директории папки с обзорами *****

         IniSectionValue [12,3]:= DirTerm;  //нет отсечения имени файла с обзором района от пути
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [12,2], IniSectionValue [12,3]);

      end;

      BookObzorTerm.Range.Copy;
      ReplaceTextTerm:='термины###';
      MyRange2 := FindInDoc(Book, ReplaceTextTerm);

         if VarIsClear(MyRange2) then begin
            ShowMessage('Текст термины### в ReplaceTextTerm НЕ найден.');
            Exit;
         end;

  MyRange2.Paste;
 end;




 //****************ВСТАВЛЯЕТ АНАЛОГИ В ОТЧЕТ WORD********************************************************************************************************


  //***       Аналоги для квартир, комнат, ЗУ, таунхаузов         ***
 //**** Процедура вставляет рисунки аналогов из документа WORD    ****
 //****                 в документ отчета                         ****
 procedure InsertAnalogi;
  var
        DirAnalogDoc, ReplaceTextAnalog : string;

 begin

    DirAnalogDoc:= DIRFile+'аналоги.docx';

    try
      BookAnalog:=WRD.Documents.Open(DirAnalogDoc);
      except
      ShowMessage('Не получилось открыть файл с аналогами .'+
      'Проверьте наличие, расположение файла');
      Exit;
    end;

    BookAnalog.Range.Copy;
    ReplaceTextAnalog:='1####1';
    MyRange2 := FindInDoc(Book, ReplaceTextAnalog);

         if VarIsClear(MyRange2) then begin
            ShowMessage('Текст 1####1 ReplaceTextAnalog НЕ найден.');
            Exit;
         end;

    MyRange2.Paste;

 end;


 //***       Аналоги для Здание с земельным участком              ***
 //**** Процедура вставляет рисунки аналогов из документа WORD    ****
 //****                 в документ отчета                         ****
 procedure InsertAnalogiBuildLand;
  var
        DirAnalogDoc, ReplaceTextAnalog : string;

 begin

    DirAnalogDoc:= DIRFile+'аналоги ЗУ.docx';

    try
      BookAnalog:=WRD.Documents.Open(DirAnalogDoc);
      except
      ShowMessage('Не получилось открыть файл с аналогами ЗУ.'+
      'Проверьте наличие, расположение файла');
      Exit;
    end;

    BookAnalog.Range.Copy;
    ReplaceTextAnalog:='1######1';
    MyRange2 := FindInDoc(Book, ReplaceTextAnalog);

         if VarIsClear(MyRange2) then begin
            ShowMessage('Текст 1######1 ReplaceTextAnalog НЕ найден.');
            Exit;
         end;

    MyRange2.Paste;                //аналоги ЗУ

    DirAnalogDoc:= DIRFile+'аналоги ДОМ+ЗУ.docx';

    try
      BookAnalogBuild:=WRD.Documents.Open(DirAnalogDoc);
      except
      ShowMessage('Не получилось открыть файл с аналогами ДОМ+ЗУ.'+
      'Проверьте наличие, расположение файла');
      Exit;
    end;

    BookAnalogBuild.Range.Copy;
    ReplaceTextAnalog:='1######2';
    MyRange2 := FindInDoc(Book, ReplaceTextAnalog);

         if VarIsClear(MyRange2) then begin
            ShowMessage('Текст 1######2 ReplaceTextAnalog НЕ найден.');
            Exit;
         end;

    MyRange2.Paste;                 //аналоги ДОМ+ЗУ
    BookAnalogBuild.Close;

 end;



 //**************ВСТАВЛЯЕТ ЗДАНИЯ АНАЛОГИ В ОТЧЕТ WORD****************************************************

//*******  Аналоги для здания из Ко Инвест справочника              *******
//******* Процедура вставляет  аналоги здания Ко Инвест             ********
//******* из документа WORD в другой документ в заданные места      ********
//******* (в данном случае это в первое место и остальное во второе ********

procedure InsertAnalogiCoInvest;

  var
        DirAnalogDoc, ReplaceTextAnalog : string;
        j                               : Integer;

 begin

    DirAnalogDoc:= DIRFile+'аналоги КО-ИНВЕСТ.docx';

    try
      BookAnalogiCoInvest:=WRD.Documents.Open(DirAnalogDoc);
      except
      ShowMessage('Не получилось открыть файл с аналогами Ко Инвест.'+
      'Проверьте наличие, расположение файла');
      Exit;
    end;
   //
    j:=BookAnalogiCoInvest.InlineShapes.Count;

     case j of
      0 : ShowMessage('В аналогах Ко Инвест ничего нет');
     end;

    if j=0  then   Exit;
   //
    BookAnalogiCoInvest.Range.Copy;

    ReplaceTextAnalog:='1#1ЗАТР';

    MyRange2 := FindInDoc(Book, ReplaceTextAnalog);

         if VarIsClear(MyRange2) then begin
            ShowMessage('Текст 1#1ЗАТР ReplaceTextAnalog НЕ найден.');
            Exit;
         end;

    MyRange2.Paste;

    BookAnalogiCoInvest.Close;

end;



//*****************ВСТАВЛЯЕТ ФОТО ОБЬЕКТА ОЦЕНКИ В ОТЧЕТ WORD*******************************************


//**** Процедура вставляет фото объекта оценки из документа WORD    ****
 //****                 в документ отчета                         ****
 procedure InsertFoto;
  var
        DirAnalogDoc, ReplaceTextAnalog : string;

 begin

    DirAnalogDoc:= DIRFile+'фото.docx';

    try
      BookFoto:=WRD.Documents.Open(DirAnalogDoc);
      except
        ShowMessage('Не получилось открыть файл с фото. Проверьте наличие файла,'+
        'расширение или что-то еще');
        Exit;
    end;

    BookFoto.Range.Copy;
    ReplaceTextAnalog:='1####2';
    MyRange2 := FindInDoc(Book, ReplaceTextAnalog);

         if VarIsClear(MyRange2) then begin
            ShowMessage('Текст 1####2 ReplaceFoto НЕ найден.');
            Exit;
         end;

    MyRange2.Paste;

 end;



//**************ВСТАВЛЯЕТ РИСУНКИ МЕСТОПОЛОЖЕНИЯ В ОТЧЕТ WORD**********************************************


//**** Процедура вставляет рисунки местоположения из документа WORD      ****
//**** в документ отчета                                                 ****
  procedure InsertLocation;
  var
       i,j,counter: Integer;
       DirLocationDoc  : string;
       ReplaceTextLoc : array [1..3] of string;

  begin
       ReplaceTextLoc [1]:='1#####1';
       ReplaceTextLoc [2]:='1#####2';
       ReplaceTextLoc [3]:='1#####3';


     DirLocationDoc:= DIRFile+'место.docx';

     try
       BookLocation:=WRD.Documents.Open(DirLocationDoc);
       except
       ShowMessage('Не получилось открыть файл с местоположением. '+
       'Проверьте наличие файла и правильность названия (фото), '+
       'путь к файлу (в той же папке где и отчет');
       end;

      counter :=0;
      counter:=BookLocation.InlineShapes.Count;
   // ShowMessage(inttostr(counter));
     case counter of
      0 : ShowMessage('В сканах местоположения нет данных');
      1 : ShowMessage('В сканах местоположения нет двух документов');
      2 : ShowMessage('В сканах местоположения не хватает одного документа');
     end;
   if counter=0  then   Exit;



       for i:=1 to counter  do
       begin

             ShpWrd3:= BookLocation.Range.InlineShapes;
             ShpWrd3.Item(i).select;
             WRD.selection.Copyaspicture;

        //   Sleep(1000);
         MyRange2 := FindInDoc(Book, ReplaceTextLoc [i]);

         if VarIsClear(MyRange2) then begin
            ShowMessage('Текст 1####№ Location Превышение по количеству скринов.');
            Exit;
         end;

        // Sleep(1000);
         MyRange2.Paste;
       end;

  end;


 //**** Запуск макроса на слияние документа с таблицей EXCEL    ************************************

   procedure MacroSli;

   begin
      Book.Activate;
      WRD.Run('Слияние');

   end;


    procedure MacroSliExpert;

   begin
      Book.Activate;
      WRD.Run('СлияниеЭксперт');

   end;

   procedure MacroSliDoc;

   begin
      Book.Activate;
      WRD.Run('ДоговорСлияниеПдф');

   end;


   procedure MacroSliCalc;

   begin
      Book.Activate;
      WRD.Run('СчетСлияниеПдф');

   end;


   procedure ExeMacros;

    begin
     // MyBook.Activate;
      EXC.Run('ТаблКакРис');

    end;

   procedure ExeMacrosZad;

    begin
     // MyBook.Activate;
      EXC.Run('ЗадАкт');

    end;

 //**** Процедура закрывает документы WORD кроме отчета**************************************************
   procedure CloseWordDocs;

   begin
    BookLocation.Close;
    BookAnalog.Close;
    BookObzor.Close;
    BookObzorObj.Close;
    BookObzorRF.Close;
    BookObzorRegion.Close;
    //ShowMessage('BookObzorRegion.Close;');
    BookObzorDoci.Close;
    //ShowMessage('BookObzorDoci.Close;');
    BookObzorRec.Close;

    //ShowMessage('BookObzorDoci.Close;');
    BookFoto.Close;

    //ShowMessage('BookFoto.Close();');
    //BookPicture.Close;    Закрываем в самой подпрограмме, так как
    //                      если закрывать здесь то начинает выдавать
    //                      ошибку, по ходу Эмбаркадеро фиксит то, что
    //                      Экзель закрыт а картинки берутся оттуда
    //                      ну типа ссылка пропала. Если нет образения к
    //                      стороннему обьекту то этого говна нет.
   end;

 //**** Процедура закрывает документ EXCEL
   procedure CloseExcel;

   begin
    EXC.DisplayAlerts := False;
    EXC.WorkBooks.Close;        //
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

procedure TForm1.AuthorClick(Sender: TObject);
begin
ShowMessage('Автор проги Пичукан Айвар... В общем то неплохой программист.');
end;

procedure TForm1.WhatDoClick(Sender: TObject);
begin
ShowMessage(' Работает когда нажмешь кнопку');
end;


//********** Нажатие кнопки Start Program (квартиры, комнаты, ЗУ, таунхаузы *********

procedure TForm1.Button2Click(Sender: TObject);
begin
    ProgBar:=0;
    ProgressBar1.Position := ProgBar ;
     Label2C:='программа начала работу';
     Form1.Label2.Caption:=Label2C;

    StartWORD;

    ProgBar:=7;
    ProgressBar1.Position := ProgBar ;
    Label2C:='приложение WORD стартовало';
    Form1.Label2.Caption:=Label2C;

   if not OpenDialog1.Execute then Exit;  //тут открывается диалог выбора файла, и если пользователь нажал "Cancel", то выходим

    WRD.visible:=true;     //делаем Ворд видимым or not
    Book:= WRD.Documents.Open(OpenDialog1.FileName);
    MyRange2:=Book.Range;
  // WRD.visible:=false;     //делаем Ворд видимым or not

    DIRFName := OpenDialog1.FileName;    //Показывает директорию открытого файла
  //  ShowMessage(DIRFName);             // с именем и расширением


    DIRName:= DIRDetect;    // определяется директория с названием файла
 //   ShowMessage(DIRName);   // без расширения  но с точкой


  //  ShowMessage(DIRFileDetect);    // определяется директория папки без названия файла
    DIRFile:= DIRFileDetect;       //и присваивается переменной

      ProgBar:=10;
    ProgressBar1.Position := ProgBar ;
     Label2C:='документ WORD по выбранной директории открыт';
     Form1.Label2.Caption:=Label2C;

    StartExcel;

    Label2C:='приложение EXCEL стартовало';
    Form1.Label2.Caption:=Label2C;

    ActExcelOpenDoc;             //Запускаем и открываем EXCEL по заданной директории

     ProgBar:=12;
    ProgressBar1.Position := ProgBar ;
     Label2C:='документ EXCEL по выбранной директории открыт';
     Form1.Label2.Caption:=Label2C;

    ExeMacros;

    Label2C:='макрос EXCEL табл как рис выполнен';
    Form1.Label2.Caption:=Label2C;
    EXC.visible:=true;       //делаем Excel видимым
 // EXC.visible:=false;      //делаем Excel невидимым

      ProgBar:=17;
    ProgressBar1.Position := ProgBar ;

    //TableAsPicturePaste(Book);
    TableAsPicturePaste(Book);

     ProgBar:=33;
    ProgressBar1.Position := ProgBar ;
    Label2C:='таблицы EXCEL вставлены, начинаю вставлять';
    Form1.Label2.Caption:=Label2C;


    //InsertPictureWord(Book);
      InsertPictureWord();

     ProgBar:=40;
    ProgressBar1.Position := ProgBar ;
    Label2C:='акт, доки, задание вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

    //BookPicture.Close;

    InsertObzor() ;

      ProgBar:=45;
    ProgressBar1.Position := ProgBar ;
    Label2C:='обзор района вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

     InsertObzorObj() ;

      ProgBar:=51;
    ProgressBar1.Position := ProgBar ;
    Label2C:='обзор obj вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

     InsertObzorRFandRegion();

     ProgBar:=56;
    ProgressBar1.Position := ProgBar ;
    Label2C:='обзор RF и Region вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

     InsertFco();

     ProgBar:=60;
    ProgressBar1.Position := ProgBar ;
    Label2C:='обзор ФСО вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

    BookObzorFco.Close();


   InsertOgr();

     ProgBar:=63;
    ProgressBar1.Position := ProgBar ;
    Label2C:='ограничения и предположения вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

    BookObzorOgr.Close();


    InsertDop();

     ProgBar:=66;
    ProgressBar1.Position := ProgBar ;
    Label2C:='допущения вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

    BookObzorDop.Close();


    InsertTerm();

     ProgBar:=68;
    ProgressBar1.Position := ProgBar ;
    Label2C:='термины и определения вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

    BookObzorTerm.Close();


    InsertLit();

     ProgBar:=69;
    ProgressBar1.Position := ProgBar ;
    Label2C:='литература вставлена в WORD документ';
    Form1.Label2.Caption:=Label2C;

    BookObzorLit.Close();




    InsertDociOcen();

    ProgBar:=71;
    ProgressBar1.Position := ProgBar ;
    Label2C:='доки оценочные вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

     InsertDociRec();

    ProgBar:=73;
    ProgressBar1.Position := ProgBar ;
    Label2C:='реквизиты оценочные вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

     CloseExcel;

     ProgBar:=75;
    ProgressBar1.Position := ProgBar ;
    Label2C:='закрыт EXCEL, начинаю вставлять аналоги';
    Form1.Label2.Caption:=Label2C;

    InsertAnalogi;

      ProgBar:=79;
    ProgressBar1.Position := ProgBar ;
    Label2C:='аналоги вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

    InsertFoto;

      ProgBar:=82;
    ProgressBar1.Position := ProgBar ;
    Label2C:='фотки вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

    InsertLocation;

     ProgBar:=85;
    ProgressBar1.Position := ProgBar ;
    Label2C:='картинки места вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

    MacroSli;

     ProgBar:=90;
    ProgressBar1.Position := ProgBar ;
     Label2C:='слияние WORD и EXCEL выполнено';
    Form1.Label2.Caption:=Label2C;

    CloseWordDocs;

     ProgBar:=95;
    ProgressBar1.Position := ProgBar ;
     Label2C:='закрыты лишние документы WORD';
    Form1.Label2.Caption:=Label2C;


  WRD.visible:=true;     //делаем Ворд видимым

   ProgBar:=100;
    ProgressBar1.Position := ProgBar ;
    Label2C:='программа '+ProgramName+' завершена';
    Form1.Label2.Caption:=Label2C;

  ShowMessage('program '+ ProgramName + ' end');
  //Book.Range(1,10).Paste;

     ProgBar:=0;
    ProgressBar1.Position := ProgBar ;
     Label2C:='программа ожидает запуска';
    Form1.Label2.Caption:=Label2C;

end;

//********* Нажатие кнопки Экспертное Заключение ************************

   procedure TForm1.Button3Click(Sender: TObject);
begin

     ProgBar:=0;
     ProgressBar1.Position := ProgBar ;
     Label2C:='программа экспертного заключения начала работу';
     Form1.Label2.Caption:=Label2C;

    StartWORD;

    ProgBar:=7;
    ProgressBar1.Position := ProgBar ;
    Label2C:='приложение WORD стартовало';
    Form1.Label2.Caption:=Label2C;

   if not OpenDialog1.Execute then Exit;  //тут открывается диалог выбора файла, и если пользователь нажал "Cancel", то выходим

    WRD.visible:=true;     //делаем Ворд видимым or not
    Book:= WRD.Documents.Open(OpenDialog1.FileName);
    MyRange2:=Book.Range;
  // WRD.visible:=false;     //делаем Ворд видимым or not

    DIRFName := OpenDialog1.FileName;    //Показывает директорию открытого файла
  //  ShowMessage(DIRFName);             // с именем и расширением


    DIRName:= DIRDetect;    // определяется директория с названием файла
 //   ShowMessage(DIRName);   // без расширения  но с точкой


  //  ShowMessage(DIRFileDetect);    // определяется директория папки без названия файла
    DIRFile:= DIRFileDetect;       //и присваивается переменной

      ProgBar:=25;
    ProgressBar1.Position := ProgBar ;
     Label2C:='документ WORD по выбранной директории открыт';
     Form1.Label2.Caption:=Label2C;

  //  StartExcel;

 //   Label2C:='приложение EXCEL стартовало';
 //   Form1.Label2.Caption:=Label2C;

    DetectExcelOpenExpertDoc;             //Определяем имя файла EXCEL по заданной директории
                                          //Файл EXCEL не открываем
     ProgBar:=45;
    ProgressBar1.Position := ProgBar ;
     Label2C:='документ EXCEL по выбранной директории выявлен';
     Form1.Label2.Caption:=Label2C;

      MacroSliExpert;

     ProgBar:=75;
    ProgressBar1.Position := ProgBar ;
     Label2C:='сливание WORD и EXCEL выполнено';
    Form1.Label2.Caption:=Label2C;


     ProgBar:=85;
    ProgressBar1.Position := ProgBar ;
     Label2C:='все сделано и слито';
    Form1.Label2.Caption:=Label2C;


  WRD.visible:=true;     //делаем Ворд видимым

   ProgBar:=100;
    ProgressBar1.Position := ProgBar ;
    Label2C:='программа '+ProgramName+' завершена (экспертное слито)';
    Form1.Label2.Caption:=Label2C;

  ShowMessage('program '+ ProgramName + ' end');


     ProgBar:=0;
    ProgressBar1.Position := ProgBar ;
     Label2C:='программа ожидает запуска';
    Form1.Label2.Caption:=Label2C;



end;

//********* Нажатие кнопки Акт осмотра и Задание на оценку *******************

procedure TForm1.Button4Click(Sender: TObject);
begin
   ProgBar:=0;
     ProgressBar1.Position := ProgBar ;
     Label2C:='программа составления Акта и Задания начала работу';
     Form1.Label2.Caption:=Label2C;



     StartExcel;

     if not OpenDialog1.Execute then Exit;
     DIRFName := OpenDialog1.FileName;    //Показывает директорию открытого файла
  //  ShowMessage(DIRFName);             // с именем и расширением
     MyBook:= EXC.WorkBooks.Open(DIRFName);


    StartWORD;

    ProgBar:=15;
    ProgressBar1.Position := ProgBar ;
    Label2C:='приложение WORD и ЭКЗЕЛЬ стартовало';
    Form1.Label2.Caption:=Label2C;


  //  ShowMessage(DIRFileDetect);    // определяется директория папки без названия файла
    DIRFile:= DIRFileDetect;       //и присваивается переменной

      ProgBar:=40;
    ProgressBar1.Position := ProgBar ;
     Label2C:='документ WORD по выбранной директории открыт';
     Form1.Label2.Caption:=Label2C;

   ExeMacrosZad;

   InsertPictureWordZad;

   CloseExcel;

  WRD.visible:=true;     //делаем Ворд видимым

   ProgBar:=100;
    ProgressBar1.Position := ProgBar ;
    Label2C:='программа '+ProgramName+' завершена (задание и акт вставлены)';
    Form1.Label2.Caption:=Label2C;

  ShowMessage('program '+ ProgramName + ' end');

     ProgBar:=0;
    ProgressBar1.Position := ProgBar ;
     Label2C:='программа ожидает запуска';
    Form1.Label2.Caption:=Label2C;


end;




//********* Нажатие кнопки Нежилое помещение *******************

procedure TForm1.Button5Click(Sender: TObject);
begin

  ProgBar:=0;
    ProgressBar1.Position := ProgBar ;
     Label2C:='программа начала работу';
     Form1.Label2.Caption:=Label2C;

    StartWORD;

    ProgBar:=7;
    ProgressBar1.Position := ProgBar ;
    Label2C:='приложение WORD стартовало';
    Form1.Label2.Caption:=Label2C;

   if not OpenDialog1.Execute then Exit;  //тут открывается диалог выбора файла, и если пользователь нажал "Cancel", то выходим

    WRD.visible:=true;     //делаем Ворд видимым or not
    Book:= WRD.Documents.Open(OpenDialog1.FileName);
    MyRange2:=Book.Range;
  // WRD.visible:=false;     //делаем Ворд видимым or not

    DIRFName := OpenDialog1.FileName;    //Показывает директорию открытого файла
  //  ShowMessage(DIRFName);             // с именем и расширением


    DIRName:= DIRDetect;    // определяется директория с названием файла
 //   ShowMessage(DIRName);   // без расширения  но с точкой


  //  ShowMessage(DIRFileDetect);    // определяется директория папки без названия файла
    DIRFile:= DIRFileDetect;       //и присваивается переменной

      ProgBar:=10;
    ProgressBar1.Position := ProgBar ;
     Label2C:='документ WORD по выбранной директории открыт';
     Form1.Label2.Caption:=Label2C;

    StartExcel;

    Label2C:='приложение EXCEL стартовало';
    Form1.Label2.Caption:=Label2C;

    ActExcelOpenDoc;             //Запускаем и открываем EXCEL по заданной директории

     ProgBar:=12;
    ProgressBar1.Position := ProgBar ;
     Label2C:='документ EXCEL по выбранной директории открыт';
     Form1.Label2.Caption:=Label2C;

 {   ExeMacros;

    Label2C:='макрос EXCEL табл как рис выполнен';
    Form1.Label2.Caption:=Label2C;
    EXC.visible:=true;       //делаем Excel видимым
 // EXC.visible:=false;      //делаем Excel невидимым

 }
      ProgBar:=17;
    ProgressBar1.Position := ProgBar ;

 {   TableAsPicturePasteBuildLand(Book);    //*******************************
  }
    ProgBar:=33;
    ProgressBar1.Position := ProgBar ;
  {  Label2C:='таблицы EXCEL вставлены, начинаю вставлять';
    Form1.Label2.Caption:=Label2C;
  }

    //InsertPictureWord(Book);
    InsertPictureWord();

     ProgBar:=45;
    ProgressBar1.Position := ProgBar ;
    Label2C:='акт, доки, задание вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

    InsertObzor() ;

      ProgBar:=50;
    ProgressBar1.Position := ProgBar ;
    Label2C:='обзор района вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

     InsertObzorObj() ;

      ProgBar:=58;
    ProgressBar1.Position := ProgBar ;
    Label2C:='обзор obj вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

    InsertObzorRFandRegion();

     ProgBar:=62;
    ProgressBar1.Position := ProgBar ;
    Label2C:='обзор RF и Region вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

    InsertDociOcen();

    ProgBar:=69;
    ProgressBar1.Position := ProgBar ;
    Label2C:='доки оценочные вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

     InsertDociRec();

    ProgBar:=71;
    ProgressBar1.Position := ProgBar ;
    Label2C:='реквизиты оценочные вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

     CloseExcel;

     ProgBar:=73;
    ProgressBar1.Position := ProgBar ;
    Label2C:='закрыт EXCEL, начинаю вставлять аналоги';
    Form1.Label2.Caption:=Label2C;

  {  InsertAnalogiBuildLand;
  }
   { InsertAnalogiCoInvest;
    }
      ProgBar:=75;
    ProgressBar1.Position := ProgBar ;
 {   Label2C:='аналоги вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;
  }
    InsertFoto;

      ProgBar:=79;
    ProgressBar1.Position := ProgBar ;
    Label2C:='фотки вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

    InsertLocation;

     ProgBar:=85;
    ProgressBar1.Position := ProgBar ;
    Label2C:='картинки места вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

    MacroSli;

     ProgBar:=90;
    ProgressBar1.Position := ProgBar ;
     Label2C:='слияние WORD и EXCEL выполнено';
    Form1.Label2.Caption:=Label2C;

  //  CloseWordDocs;

     ProgBar:=95;
    ProgressBar1.Position := ProgBar ;
     Label2C:='закрыты лишние документы WORD';
    Form1.Label2.Caption:=Label2C;


  WRD.visible:=true;     //делаем Ворд видимым

   ProgBar:=100;
    ProgressBar1.Position := ProgBar ;
    Label2C:='программа '+ProgramName+' завершена';
    Form1.Label2.Caption:=Label2C;

  ShowMessage('program '+ ProgramName + ' end');
  //Book.Range(1,10).Paste;

     ProgBar:=0;
    ProgressBar1.Position := ProgBar ;
     Label2C:='программа ожидает запуска';
    Form1.Label2.Caption:=Label2C;

end;





//********* Нажатие кнопки Здание с земельным участком *******************

procedure TForm1.Button1Click(Sender: TObject);
begin
    ProgBar:=0;
    ProgressBar1.Position := ProgBar ;
     Label2C:='программа начала работу';
     Form1.Label2.Caption:=Label2C;

    StartWORD;

    ProgBar:=7;
    ProgressBar1.Position := ProgBar ;
    Label2C:='приложение WORD стартовало';
    Form1.Label2.Caption:=Label2C;

   if not OpenDialog1.Execute then Exit;  //тут открывается диалог выбора файла, и если пользователь нажал "Cancel", то выходим

    WRD.visible:=true;     //делаем Ворд видимым or not
    Book:= WRD.Documents.Open(OpenDialog1.FileName);
    MyRange2:=Book.Range;
  // WRD.visible:=false;     //делаем Ворд видимым or not

    DIRFName := OpenDialog1.FileName;    //Показывает директорию открытого файла
  //  ShowMessage(DIRFName);             // с именем и расширением


    DIRName:= DIRDetect;    // определяется директория с названием файла
 //   ShowMessage(DIRName);   // без расширения  но с точкой


  //  ShowMessage(DIRFileDetect);    // определяется директория папки без названия файла
    DIRFile:= DIRFileDetect;       //и присваивается переменной

      ProgBar:=10;
    ProgressBar1.Position := ProgBar ;
     Label2C:='документ WORD по выбранной директории открыт';
     Form1.Label2.Caption:=Label2C;

    StartExcel;

    Label2C:='приложение EXCEL стартовало';
    Form1.Label2.Caption:=Label2C;

    ActExcelOpenDoc;             //Запускаем и открываем EXCEL по заданной директории

     ProgBar:=12;
    ProgressBar1.Position := ProgBar ;
     Label2C:='документ EXCEL по выбранной директории открыт';
     Form1.Label2.Caption:=Label2C;

    ExeMacros;

    Label2C:='макрос EXCEL табл как рис выполнен';
    Form1.Label2.Caption:=Label2C;
    EXC.visible:=true;       //делаем Excel видимым
 // EXC.visible:=false;      //делаем Excel невидимым

      ProgBar:=17;
    ProgressBar1.Position := ProgBar ;

    TableAsPicturePasteBuildLand(Book);

     ProgBar:=33;
    ProgressBar1.Position := ProgBar ;
    Label2C:='таблицы EXCEL вставлены, начинаю вставлять';
    Form1.Label2.Caption:=Label2C;


    //InsertPictureWord(Book);
    InsertPictureWord();

     ProgBar:=48;
    ProgressBar1.Position := ProgBar ;
    Label2C:='акт, доки, задание вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

    InsertObzor() ;

      ProgBar:=55;
    ProgressBar1.Position := ProgBar ;
    Label2C:='обзор района вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

     InsertObzorObj() ;

      ProgBar:=58;
    ProgressBar1.Position := ProgBar ;
    Label2C:='обзор obj вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

    InsertObzorRFandRegion();

     ProgBar:=63;
    ProgressBar1.Position := ProgBar ;
    Label2C:='обзор RF и Region вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

    InsertDociOcen();

    ProgBar:=69;
    ProgressBar1.Position := ProgBar ;
    Label2C:='доки оценочные вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

     InsertDociRec();

    ProgBar:=71;
    ProgressBar1.Position := ProgBar ;
    Label2C:='реквизиты оценочные вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

     CloseExcel;

     ProgBar:=73;
    ProgressBar1.Position := ProgBar ;
    Label2C:='закрыт EXCEL, начинаю вставлять аналоги';
    Form1.Label2.Caption:=Label2C;

    InsertAnalogiBuildLand;

    InsertAnalogiCoInvest;

      ProgBar:=75;
    ProgressBar1.Position := ProgBar ;
    Label2C:='аналоги вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

    InsertFoto;

      ProgBar:=79;
    ProgressBar1.Position := ProgBar ;
    Label2C:='фотки вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

    InsertLocation;

     ProgBar:=85;
    ProgressBar1.Position := ProgBar ;
    Label2C:='картинки места вставлены в WORD документ';
    Form1.Label2.Caption:=Label2C;

    MacroSli;

     ProgBar:=90;
    ProgressBar1.Position := ProgBar ;
     Label2C:='слияние WORD и EXCEL выполнено';
    Form1.Label2.Caption:=Label2C;

    CloseWordDocs;

     ProgBar:=95;
    ProgressBar1.Position := ProgBar ;
     Label2C:='закрыты лишние документы WORD';
    Form1.Label2.Caption:=Label2C;


  WRD.visible:=true;     //делаем Ворд видимым

   ProgBar:=100;
    ProgressBar1.Position := ProgBar ;
    Label2C:='программа '+ProgramName+' завершена';
    Form1.Label2.Caption:=Label2C;

  ShowMessage('program '+ ProgramName + ' end');
  //Book.Range(1,10).Paste;

     ProgBar:=0;
    ProgressBar1.Position := ProgBar ;
     Label2C:='программа ожидает запуска';
    Form1.Label2.Caption:=Label2C;

end;




//********** Процедура определения названия папки текущей
//**********       директории расположения файла

function DIRFolderDetect: string;

 var

 Smb         : Char;
 i , L , k      : Integer;
 DIRECTORY   : string;

 begin
        DIRECTORY:=' ';
        DIRECTORY:=DIRFileDetect;
        L:=Length(DIRECTORY);
        Delete(DIRECTORY, L-1, L);
        L:=Length(DIRECTORY);
        i:=L;
        k:=0;
        Smb:=DIRECTORY[i];

        while (Smb<>'\') do
          begin
             k:=k+1;
             i:=i-1;
             Smb:=DIRECTORY[i];
          end;

          //ShowMessage(DIRECTORY);

          Delete(DIRECTORY,1,L-k);
          //LengthDir:=  Length(DIRECTORY);

          //ShowMessage(DIRECTORY);
          result:=DIRECTORY;

 end;








//**********Нажатие кнопки Договор слияние *****************

procedure TForm1.Button6Click(Sender: TObject);
begin

     ProgBar:=0;
     ProgressBar1.Position := ProgBar ;
     Label2C:='программа слияния договора начала работу';
     Form1.Label2.Caption:=Label2C;

    StartWORD;

    ProgBar:=7;
    ProgressBar1.Position := ProgBar ;
    Label2C:='приложение WORD стартовало';
    Form1.Label2.Caption:=Label2C;

   if not OpenDialog1.Execute then Exit;  //тут открывается диалог выбора файла, и если пользователь нажал "Cancel", то выходим

    WRD.visible:=true;     //делаем Ворд видимым or not
    Book:= WRD.Documents.Open(OpenDialog1.FileName);
    MyRange2:=Book.Range;
  // WRD.visible:=false;     //делаем Ворд видимым or not

    DIRFName := OpenDialog1.FileName;    //Показывает директорию открытого файла
  //  ShowMessage(DIRFName);             // с именем и расширением


  //  DIRName:= DIRDetect;    // определяется директория с названием файла
 //   ShowMessage(DIRName);   // без расширения  но с точкой






    //ShowMessage(DIRFileDetect);    // определяется директория папки без названия файла

    DIRFile:= DIRFolderDetect;       //и присваивается переменной

    //  ShowMessage(DIRFile);
  //    ShowMessage(DIRFname);
      ProgBar:=25;
    ProgressBar1.Position := ProgBar ;
     Label2C:='документ WORD по выбранной директории открыт';
     Form1.Label2.Caption:=Label2C;

    StartExcel;

    Label2C:='приложение EXCEL стартовало';
    Form1.Label2.Caption:=Label2C;






    //DetectExcelOpenExpertDoc;             //Определяем имя файла EXCEL по заданной директории
                                          //Файл EXCEL не открываем
     ProgBar:=45;
    ProgressBar1.Position := ProgBar ;
     Label2C:='документ EXCEL по выбранной директории выявлен';
     Form1.Label2.Caption:=Label2C;

      MacroSliDoc;

     ProgBar:=75;
    ProgressBar1.Position := ProgBar ;
     Label2C:='сливание WORD и EXCEL выполнено';
    Form1.Label2.Caption:=Label2C;


     ProgBar:=85;
    ProgressBar1.Position := ProgBar ;
     Label2C:='все сделано и слито';
    Form1.Label2.Caption:=Label2C;


  WRD.visible:=true;     //делаем Ворд видимым

   ProgBar:=100;
    ProgressBar1.Position := ProgBar ;
    Label2C:='программа '+ProgramName+' завершена (Договор слито)';
    Form1.Label2.Caption:=Label2C;

  ShowMessage('program '+ ProgramName + ' end');


     ProgBar:=0;
    ProgressBar1.Position := ProgBar ;
     Label2C:='программа ожидает запуска';
    Form1.Label2.Caption:=Label2C;



end;





//**********  Нажатие кнопки счет слияние *****************************

 procedure TForm1.Button7Click(Sender: TObject);
begin

  ProgBar:=0;
     ProgressBar1.Position := ProgBar ;
     Label2C:='программа слияния Счет начала работу';
     Form1.Label2.Caption:=Label2C;

    StartWORD;

    ProgBar:=7;
    ProgressBar1.Position := ProgBar ;
    Label2C:='приложение WORD стартовало';
    Form1.Label2.Caption:=Label2C;

   if not OpenDialog1.Execute then Exit;  //тут открывается диалог выбора файла, и если пользователь нажал "Cancel", то выходим

    WRD.visible:=true;     //делаем Ворд видимым or not
    Book:= WRD.Documents.Open(OpenDialog1.FileName);
    MyRange2:=Book.Range;
  // WRD.visible:=false;     //делаем Ворд видимым or not

    DIRFName := OpenDialog1.FileName;    //Показывает директорию открытого файла
  //  ShowMessage(DIRFName);             // с именем и расширением


  //  DIRName:= DIRDetect;    // определяется директория с названием файла
 //   ShowMessage(DIRName);   // без расширения  но с точкой






   // ShowMessage(DIRFileDetect);    // определяется директория папки без названия файла

    DIRFile:= DIRFolderDetect;       //и присваивается переменной

     // ShowMessage(DIRFile);
  //    ShowMessage(DIRFname);
      ProgBar:=25;
    ProgressBar1.Position := ProgBar ;
     Label2C:='документ WORD по выбранной директории открыт';
     Form1.Label2.Caption:=Label2C;

    StartExcel;

    Label2C:='приложение EXCEL стартовало';
    Form1.Label2.Caption:=Label2C;






    //DetectExcelOpenExpertDoc;             //Определяем имя файла EXCEL по заданной директории
                                          //Файл EXCEL не открываем
     ProgBar:=45;
    ProgressBar1.Position := ProgBar ;
     Label2C:='документ EXCEL по выбранной директории выявлен';
     Form1.Label2.Caption:=Label2C;

      MacroSliCalc;

     ProgBar:=75;
    ProgressBar1.Position := ProgBar ;
     Label2C:='сливание WORD и EXCEL выполнено';
    Form1.Label2.Caption:=Label2C;


     ProgBar:=85;
    ProgressBar1.Position := ProgBar ;
     Label2C:='все сделано и слито';
    Form1.Label2.Caption:=Label2C;


  WRD.visible:=true;     //делаем Ворд видимым

   ProgBar:=100;
    ProgressBar1.Position := ProgBar ;
    Label2C:='программа '+ProgramName+' завершена (Счет слито)';
    Form1.Label2.Caption:=Label2C;

  ShowMessage('program '+ ProgramName + ' end');


     ProgBar:=0;
    ProgressBar1.Position := ProgBar ;
     Label2C:='программа ожидает запуска';
    Form1.Label2.Caption:=Label2C;





end;




//*************** процедура выдачи серийника диска (скачал с инета) ******************
   function GetHardDiskSerial(const DriveLetter: Char): string;
var
  NotUsed:     DWORD;
  VolumeFlags: DWORD;
  VolumeInfo:  array[0..MAX_PATH] of Char;
  VolumeSerialNumber: DWORD;
begin
  GetVolumeInformation(PChar(DriveLetter + ':\'),
    nil, SizeOf(VolumeInfo), @VolumeSerialNumber, NotUsed,
    VolumeFlags, nil, 0);
 // Result := Format('Label = %s   VolSer = %8.8X',
 // [VolumeInfo, VolumeSerialNumber])

    Result := Format('VolSer = %8.8X',
    [VolumeSerialNumber])
end;

 //************** шифрование *********************************************************
 function EnCrypt (const InputCryptData : string) : string;
 var
   SyspendCrypt       : string;
   Syspend            : char;
   i                  : Integer;
   LengthSyspendCrypt : Integer;

   begin
   SyspendCrypt :=  InputCryptData + InputCryptData;   //плюсуем два слова типа string
   LengthSyspendCrypt := Length(SyspendCrypt);
   i := 1;

    while i < LengthSyspendCrypt do    //меняем символы между собой относительно центра слова
       begin
         Syspend := SyspendCrypt[i];
         SyspendCrypt[i] := SyspendCrypt[i+1];
         SyspendCrypt[i+1] := Syspend;
         i := i + 2;
       end;

      while i < LengthSyspendCrypt do  //меняем символы у двух половинок друг с другом
       begin
         Syspend := SyspendCrypt[i];
         SyspendCrypt[i] := SyspendCrypt[Round(LengthSyspendCrypt/2)+i];
         SyspendCrypt[Round(LengthSyspendCrypt/2)+i] := Syspend;
         i := i + 2;
       end;

     Result := SyspendCrypt;

   end;

 //*********************** Проверка подлинности программы *******************************

   procedure CheckCrypt;

   var KeyCrypt       : string;
   var KeyCryptDisc1       : string;
   var KeyCryptDisc2       : string;
   var KeyCryptDisc3       : string;
   var KeyCryptResult : Boolean;

   begin
     //******************** Проверка железа **********************************************


     KeyCryptDisc1 := GetHardDiskSerial('c');
     KeyCryptDisc2 := EnCrypt (KeyCryptDisc1);
     KeyCryptDisc3 := IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [2,2], IniSectionValue [2,3]);
     //showMessage(KeyCryptDisc3);
     if KeyCryptDisc2 = KeyCryptDisc3
      then

       Exit;

    //********* Если железо не совпало то проверка пароля ******************************


     Password := InputBox('AUTHORIZATION','Пожалуйста укажите пароль','Одуванчик') ;
    // showMessage(Password);
    // ShowMessage(EnCrypt(Password));

    KeyCrypt:='lbvegaardnlbvegaardn';

    if EnCrypt(Password) = KeyCrypt then    KeyCryptResult := True
                                    else    KeyCryptResult := False;



    if KeyCryptResult then  IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [2,2], KeyCryptDisc2)
                      else
                      begin
                        ShowMessage('Неверный пароль программы');
                        Application.Terminate;
                      end;





   end;


   //********************** Процедуры работы с INI  файлами **************************
 procedure IniFileCreate;
    begin
    IniSectionValue [1,1]:= 'FileLocation';
    IniSectionValue [1,2]:= 'LocationOBZOR';
    IniSectionValue [1,3]:= 'Z:\GRAND NEVA\2014\ОБЗОРЫ районы\'; //Как пример значения в секции INI файла
    IniFile:=TIniFile.Create(ExtractFilePath(Application.ExeName)+ProgramName+'.ini');
    //******************* For Crypt Information **********************************
    IniSectionValue [2,2]:= 'Capasity Data';
    IniSectionValue [2,3]:= 'shhhdbbdb56';
    IniSectionValue [4,2]:= 'Object OBZOR Location';
    IniSectionValue [4,3]:= 'Z:\GRAND NEVA\2014\ОБЗОРЫ районы\';
    IniSectionValue [5,2]:= 'RF OBZOR Location';
    IniSectionValue [5,3]:= 'Z:\GRAND NEVA\2014\ОБЗОР РФ\';
    IniSectionValue [6,2]:= 'REGION OBZOR Location';
    IniSectionValue [6,3]:= 'Z:\GRAND NEVA\2014\ОБЗОР регион\';
    IniSectionValue [7,2]:= 'DOCI Ocenschik Location';
    IniSectionValue [7,3]:= 'Z:\GRAND NEVA\2014\доки оценщика\';
    IniSectionValue [8,2]:= 'Recvizity Ocenschik Location';
    IniSectionValue [8,3]:= 'Z:\GRAND NEVA\2014\доки оценщика\';
    IniSectionValue [9,2]:= 'Fco Location';
    IniSectionValue [9,3]:= 'Z:\GRAND NEVA\2014\Fco\';
    IniSectionValue [10,2]:= 'Ogr Location';
    IniSectionValue [10,3]:= 'Z:\GRAND NEVA\2014\Ogr\';
    IniSectionValue [11,2]:= 'Dop Location';
    IniSectionValue [11,3]:= 'Z:\GRAND NEVA\2014\Dop\';
    IniSectionValue [12,2]:= 'Term Location';
    IniSectionValue [12,3]:= 'Z:\GRAND NEVA\2014\Term\';
    IniSectionValue [13,2]:= 'Lit Location';
    IniSectionValue [13,3]:= 'Z:\GRAND NEVA\2014\Literature\';


  end;



 //*********************** MAIN PROGRAMM *******************************************************

begin
   ProgramName:='Blevantone 9.1 dev';


   ProgBar:=0;


  IniFileCreate;

  CheckCrypt;




end.

