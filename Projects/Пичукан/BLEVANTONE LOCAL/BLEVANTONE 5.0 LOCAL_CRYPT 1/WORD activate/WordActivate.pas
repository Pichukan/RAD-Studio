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

    procedure Button2Click(Sender: TObject);
    procedure AuthorClick(Sender: TObject);
    procedure WhatDoClick(Sender: TObject);


  //  procedure Label1Click(Sender: TObject);
  //  procedure Label2Click(Sender: TObject);




  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

  WRD,Book,BookPicture,BookObzor,BookAnalog,BookLocation,BookFoto,
      BookBlevantone,RangeBookBlevantone,wdInlineShapes,wdInlineShapes2: OleVariant;
  EXC,MyBook,MyWorkSheet,MyWorkSheet2,MyRange,MyRange2,RangeObzor,
      Shp,ShpWrd,ShpWrd2,ShpWrd3,vstart,vend: OleVariant;
  var W,ObzorValue:variant;
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
  var  Label2C     : string;
  var  IniFile: TIniFile;
  var  IniSectionValue : array [1..10,1..3] of string;
  var Password    : string;
 // var Dialog : OpenDialog1;

implementation

{$R *.dfm}

//*************** �������� ������� ������������ �������� EXCEL ****************
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

//***************** �������� ���������� �� EXCEL ************************
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

//************** �������� ���������� �� WORD **********************************
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

//************* ������� ������� ���������� �� ����������� EXCEL ***********************************

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


//*********************** ������� ������� ���������� �� ����������� WORD *******************************
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

//*********************** ������� ������� ���������� EXCEL *********************
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
    ShowMessage('���������� Excel �� ���������� �� ������ ���������');
     result:=False;
    end;
 end;


//*********************** ������� ������� ���������� WORD *******************************
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
    ShowMessage('���������� WORD �� ���������� �� ������ ���������');
     result:=False;
    end;
 end;

 //**** ������������ ���������� � ��������� ����� ��� ����������  �� � ������ ****************
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

//****************** ������������ ���������� ����� ��� �������� ����� **************
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

 //****************** ������������ ���������� ����� ��� �������� �����  **************
 //******************               �� �������� ����������              **************
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


 // ***************** ������� ���� � ����� �������� ����� �� ����� ��������� ************************************************
 // *****************         � ���� �� ���� ������ ������ range             *******************************
  function FindInDoc(const aWdDoc : OleVariant; const aSearchText : String) : oleVariant;
const
  wdFindStop = 0; //��������� ����� ��� ���������� ������� ���������.
var
  wdRng, wdFind : OleVariant;
begin
  VarClear(Result);
  //��������, ������������ �� ���������� ���������.
  wdRng := aWdDoc.Content;
  //��������� ������.
  wdFind := wdRng.Find;
  //wdFind.ClearFormatting;
  wdFind.Text := aSearchText;
  //True - ����� ����� �� ������ - � ����� ���������.
  wdFind.Forward := True;
  //wdFindStop - ��������� ����� ��� ���������� ������� ���������.
  //wdFind.Wrap := wdFindStop;
  //����� ������.
  if wdFind.Execute then Result := wdRng;
end;


//******** ��������� �������� ��������� EXCEL �� ��������� ���� **************************************************************
procedure ActExcelOpenDoc;
begin
 {
  if StartExcel then
   begin
    ShowMessage('������� Excel �������');
   end;
  }
  DIRExName:=DIRName+'xlsm';
//  ShowMessage('������� Excel �����������');
  MyBook:= EXC.WorkBooks.Open(DIRExName);  //��������� ��������
//  ShowMessage(DIRExName);   //���������� ���������� ��������� ����� � ������ � �����������

end;


//********** ��������� ��������� ������ EXCEL ���������� ������� ��� �������  ***********
//********** �� ����� EXCEL � ����������� �� ��� ��, � ����� ��������� ���    ***********
//********** ������� ��������� � �������� WORD �� �������� ����� ����� ������ ***********
//********** ��������� ������                                                 ***********
procedure TableAsPicturePaste(Book : OleVariant);
  var  ReplaceText : array [1..10] of string;
                 j : Integer;

  begin
   //   ShowMessage('0');
    //  EXC.Run('PERSONAL.XLSB!����������');
    //  EXC.Run('���');
    //  EXC.Run('����������');
   //  ShowMessage('000');
   //  Sleep(1000);
     MyWorkSheet:=MyBook.Sheets['������� ����'];

   //  ShowMessage('001');
   //  Sleep(500);
     Shp := MyWorkSheet.Shapes.Item(1);
    //  MyRange:=EXC.Range['a2'];
    //   Book.Activate;
    // MyRange2:=EXC.Range['h1'];
    //  Book.Range(1,10).Paste;
    //  ShowMessage('0002');
     ReplaceText [1]:='1#1';        //���������� �� ������� ����� � WORD
     ReplaceText [2]:='1#2';        //���������� �� ������� ����� � WORD
     ReplaceText [3]:='1#3';        //���������� �� ������� ����� � WORD
     ReplaceText [4]:='1#4';        //���������� �� ������� ����� � WORD
     ReplaceText [5]:='1#5';        //���������� �� ������� ����� � WORD
     ReplaceText [6]:='1#6';        //���������� �� ������� ����� � WORD

     j:=1;

     repeat
      //ShowMessage('002');
        Shp := MyWorkSheet.Shapes.Item(j);
        Shp.Copy;
      //ShowMessage('111');
        MyRange2 := FindInDoc(Book, ReplaceText [j]);

       if VarIsClear(MyRange2) then begin
          ShowMessage('����� �� ������.');
          Exit;
       end;

         MyRange2.Select;      //�������� ��������� �����.
         MyRange2.Paste;       //�������� ��������� ����� �� ������ ������.

        j:= j+1;
     until ReplaceText [j]  = '';

   //ShowMessage('������� �� ������ ���������');

  end;

//******* ��������� ���������  ������� �� ��������� WORD   ********
//******* � ������ �������� � �������� �����               ********
//******* (� ������ ������ ��� ���. � ��� �������          ********
procedure InsertPictureWord(Book : OleVariant);
 var
 ReplaceText   : array [1..5] of string;
 ReplaceTextValue : string;
 DirPictureDoc : string;
 i,j             : Integer;

begin
  DirPictureDoc:= DIRFile+'����.docx';
  BookPicture:=WRD.Documents.Open(DirPictureDoc);


   ReplaceText [1]:='1##1';
   ReplaceText [2]:='1##2';
   ReplaceText [3]:='1##3';
   ReplaceText [4]:='';
   j:=0;

    j:=BookPicture.InlineShapes.Count;

     case j of
      0 : ShowMessage('� ����� ��� ������');
      1 : ShowMessage('� ����� ��� ���� ������� � ��������� ����������');
      2 : ShowMessage('� ����� �� ������� ������ ���������� � ���� �������');
      3 : ShowMessage('� ����� �� ������� ������ ���������� � ���� �������');
      4 : ShowMessage('� ����� ��� ������ ����������');
     end;
   if j=0  then   Exit;





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
             ShowMessage('����� � ���������;'+ReplaceTextValue+' �� ������.');
             Exit;
         end;


            MyRange2.Paste;

 //*************************************
         if i=1 then
            begin
           ProgBar:=43;
           Form1.ProgressBar1.Position := ProgBar ;
           Label2C:='������� �������, ����� ��� ������� � ��������';
           Form1.Label2.Caption:=Label2C;
          end;

         if i=4 then
           begin
           ProgBar:=50;
           Form1.ProgressBar1.Position := ProgBar ;
           Label2C:='������� ���, ����� ��������� ������ �� ����';
           Form1.Label2.Caption:=Label2C;
          end;

         if i=12 then
            begin
             ProgBar:=54;
             Form1.ProgressBar1.Position := ProgBar ;
             Label2C:='�� ��������, �������, �����...';
             Form1.Label2.Caption:=Label2C;
            end;

 //**********************************************

         if (i<j) xor (i<>1) xor (i<>4) then  MyRange2.InsertAfter(ReplaceTextValue)  ;

       end;




end;

//**** ��������� ���� �� ������� EXCEL �������� ������               *****
//**** ��������� �������� WORD � ������� �� �������� ���������       *****
//**** � ��������� � �������� WORD ������                            *****
 procedure InsertObzor();
 var
    DirObzor,ReplaceTextObzor : string;
    DefaultReadIniFile        : string;

 begin
      MyWorkSheet2:=MyBook.Sheets['����'];
      RangeObzor:=MyWorkSheet2.Range['b22'];
      ObzorValue:= RangeObzor.Value;

    //ShowMessage(vartostr(ObzorValue));

        DefaultReadIniFile := 'Z:\GRAND NEVA\2014\������ ������Z\';   //�������� ���������� ������ �� ���������
        DirObzor:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [1,2], DefaultReadIniFile)+ObzorValue+' �����.docx';

      try
      BookObzor:=WRD.Documents.Open(DirObzor);
      except

         ShowMessage('BLEVANTONE �� ���� ������� ���� � ��������.'+
       ' ���������� ������� ������������ ��������� � ������� ������ '+
        ObzorValue+' �� ���, ���� ���, ������ ������� ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //��� ����������� ������ ������ �����, � ���� ������������ ����� "Cancel", �� �������

          DirObzor := Form1.OpenDialog1.FileName;
          BookObzor:=WRD.Documents.Open(DirObzor);

         //******* ����������� � ������ � INI ���� ���������� ����� � �������� *****

         IniSectionValue [1,3]:= DIRFileDetect2(DirObzor);  //��������� ����� ����� � ������� ������ �� ����
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [1,2], IniSectionValue [1,3]);

      end;

      BookObzor.Range.Copy;
      ReplaceTextObzor:='1###1';
      MyRange2 := FindInDoc(Book, ReplaceTextObzor);

         if VarIsClear(MyRange2) then begin
            ShowMessage('����� 1###1 ReplaceTextObzor �� ������.');
            Exit;
         end;

  MyRange2.Paste;
 end;

 //**** ��������� ��������� ������� �������� �� ��������� WORD    ****
 //****                 � �������� ������                         ****
 procedure InsertAnalogi;
  var
        DirAnalogDoc, ReplaceTextAnalog : string;

 begin

    DirAnalogDoc:= DIRFile+'�������.docx';

    try
      BookAnalog:=WRD.Documents.Open(DirAnalogDoc);
      except
      ShowMessage('�� ���������� ������� ���� � ��������� .'+
      '��������� �������, ������������ �����');
      Exit;
    end;

    BookAnalog.Range.Copy;
    ReplaceTextAnalog:='1####1';
    MyRange2 := FindInDoc(Book, ReplaceTextAnalog);

         if VarIsClear(MyRange2) then begin
            ShowMessage('����� 1####1 ReplaceTextAnalog �� ������.');
            Exit;
         end;

    MyRange2.Paste;

 end;



//**** ��������� ��������� ���� ������� ������ �� ��������� WORD    ****
 //****                 � �������� ������                         ****
 procedure InsertFoto;
  var
        DirAnalogDoc, ReplaceTextAnalog : string;

 begin

    DirAnalogDoc:= DIRFile+'����.docx';

    try
      BookFoto:=WRD.Documents.Open(DirAnalogDoc);
      except
        ShowMessage('�� ���������� ������� ���� � ����. ��������� ������� �����,'+
        '���������� ��� ���-�� ���');
        Exit;
    end;

    BookFoto.Range.Copy;
    ReplaceTextAnalog:='1####2';
    MyRange2 := FindInDoc(Book, ReplaceTextAnalog);

         if VarIsClear(MyRange2) then begin
            ShowMessage('����� 1####2 ReplaceFoto �� ������.');
            Exit;
         end;

    MyRange2.Paste;

 end;



//**** ��������� ��������� ������� �������������� �� ��������� WORD      ****
//**** � �������� ������                                                 ****
  procedure InsertLocation;
  var
       i,j,counter: Integer;
       DirLocationDoc  : string;
       ReplaceTextLoc : array [1..3] of string;

  begin
       ReplaceTextLoc [1]:='1#####1';
       ReplaceTextLoc [2]:='1#####2';
       ReplaceTextLoc [3]:='1#####3';


     DirLocationDoc:= DIRFile+'�����.docx';

     try
       BookLocation:=WRD.Documents.Open(DirLocationDoc);
       except
       ShowMessage('�� ���������� ������� ���� � ���������������. '+
       '��������� ������� ����� � ������������ �������� (����), '+
       '���� � ����� (� ��� �� ����� ��� � �����');
       end;

      counter :=0;
      counter:=BookLocation.InlineShapes.Count;
   // ShowMessage(inttostr(counter));
     case counter of
      0 : ShowMessage('� ������ �������������� ��� ������');
      1 : ShowMessage('� ������ �������������� ��� ���� ����������');
      2 : ShowMessage('� ������ �������������� �� ������� ������ ���������');
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
            ShowMessage('����� 1####� Location �� ������.');
            Exit;
         end;

        // Sleep(1000);
         MyRange2.Paste;
       end;

  end;


 //**** ������ ������� �� ������� ��������� � �������� EXCEL    ****

   procedure MacroSli;

   begin
      Book.Activate;
      WRD.Run('�������');

   end;


   procedure ExeMacros;

    begin
     // MyBook.Activate;
      EXC.Run('����������');

    end;

 //**** ��������� ��������� ��������� WORD ����� ������
   procedure CloseWordDocs;

   begin
    BookLocation.Close;
    BookAnalog.Close;
    BookObzor.Close;
    BookPicture.Close;
    BookFoto.Close;
   end;

 //**** ��������� ��������� �������� EXCEL
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

procedure TForm1.AuthorClick(Sender: TObject);
begin
ShowMessage('����� ����� ������� �����... � ����� �� �������� �����������.');
end;

procedure TForm1.WhatDoClick(Sender: TObject);
begin
ShowMessage(' �������� ����� ������� ������');
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
    ProgBar:=0;
    ProgressBar1.Position := ProgBar ;
     Label2C:='��������� ������ ������';
     Form1.Label2.Caption:=Label2C;

    StartWORD;

    ProgBar:=7;
    ProgressBar1.Position := ProgBar ;
    Label2C:='���������� WORD ����������';
    Form1.Label2.Caption:=Label2C;

   if not OpenDialog1.Execute then Exit;  //��� ����������� ������ ������ �����, � ���� ������������ ����� "Cancel", �� �������

    WRD.visible:=true;     //������ ���� ������� or not
    Book:= WRD.Documents.Open(OpenDialog1.FileName);
    MyRange2:=Book.Range;
  // WRD.visible:=false;     //������ ���� ������� or not

    DIRFName := OpenDialog1.FileName;    //���������� ���������� ��������� �����
  //  ShowMessage(DIRFName);             // � ������ � �����������


    DIRName:= DIRDetect;    // ������������ ���������� � ��������� �����
 //   ShowMessage(DIRName);   // ��� ����������  �� � ������


  //  ShowMessage(DIRFileDetect);    // ������������ ���������� ����� ��� �������� �����
    DIRFile:= DIRFileDetect;       //� ������������� ����������

      ProgBar:=10;
    ProgressBar1.Position := ProgBar ;
     Label2C:='�������� WORD �� ��������� ���������� ������';
     Form1.Label2.Caption:=Label2C;

    StartExcel;

    Label2C:='���������� EXCEL ����������';
    Form1.Label2.Caption:=Label2C;

    ActExcelOpenDoc;             //��������� � ��������� EXCEL �� �������� ����������

     ProgBar:=12;
    ProgressBar1.Position := ProgBar ;
     Label2C:='�������� EXCEL �� ��������� ���������� ������';
     Form1.Label2.Caption:=Label2C;

    ExeMacros;

    Label2C:='������ EXCEL ���� ��� ��� ��������';
    Form1.Label2.Caption:=Label2C;
    EXC.visible:=true;       //������ Excel �������
 // EXC.visible:=false;      //������ Excel ���������

      ProgBar:=17;
    ProgressBar1.Position := ProgBar ;

    TableAsPicturePaste(Book);

     ProgBar:=33;
    ProgressBar1.Position := ProgBar ;
    Label2C:='������� EXCEL ���������, ������� ���������';
    Form1.Label2.Caption:=Label2C;


    InsertPictureWord(Book);

     ProgBar:=58;
    ProgressBar1.Position := ProgBar ;
    Label2C:='���, ����, ������� ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

    InsertObzor() ;

      ProgBar:=65;
    ProgressBar1.Position := ProgBar ;
    Label2C:='����� ������ ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

     CloseExcel;

     ProgBar:=69;
    ProgressBar1.Position := ProgBar ;
    Label2C:='������ EXCEL, ������� ��������� �������';
    Form1.Label2.Caption:=Label2C;

    InsertAnalogi;

      ProgBar:=75;
    ProgressBar1.Position := ProgBar ;
    Label2C:='������� ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

    InsertFoto;

      ProgBar:=79;
    ProgressBar1.Position := ProgBar ;
    Label2C:='����� ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

    InsertLocation;

     ProgBar:=85;
    ProgressBar1.Position := ProgBar ;
    Label2C:='�������� ����� ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

    MacroSli;

     ProgBar:=90;
    ProgressBar1.Position := ProgBar ;
     Label2C:='������� WORD � EXCEL ���������';
    Form1.Label2.Caption:=Label2C;

    CloseWordDocs;

     ProgBar:=95;
    ProgressBar1.Position := ProgBar ;
     Label2C:='������� ������ ��������� WORD';
    Form1.Label2.Caption:=Label2C;


  WRD.visible:=true;     //������ ���� �������

   ProgBar:=100;
    ProgressBar1.Position := ProgBar ;
    Label2C:='��������� '+ProgramName+' ���������';
    Form1.Label2.Caption:=Label2C;

  ShowMessage('program '+ ProgramName + ' end');
  //Book.Range(1,10).Paste;

     ProgBar:=0;
    ProgressBar1.Position := ProgBar ;
     Label2C:='��������� ������� �������';
    Form1.Label2.Caption:=Label2C;

end;
 //********************** ��������� ������ � INI  ������� **************************
 procedure IniFileCreate;
    begin
    IniSectionValue [1,1]:= 'FileLocation';
    IniSectionValue [1,2]:= 'LocationOBZOR';
    IniSectionValue [1,3]:= 'Z:\GRAND NEVA\2014\������ ������\'; //��� ������ �������� � ������ INI �����
    IniFile:=TIniFile.Create(ExtractFilePath(Application.ExeName)+ProgramName+'.ini');

  end;

 //*************** ��������� ������ ��������� ����� (������ � �����) ******************
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

 //************** ���������� *********************************************************
 function EnCrypt (const InputCryptData : string) : string;
 var
   SyspendCrypt       : string;
   Syspend            : char;
   i                  : Integer;
   LengthSyspendCrypt : Integer;

   begin
   SyspendCrypt :=  InputCryptData + InputCryptData;   //������� ��� ����� ���� string
   LengthSyspendCrypt := Length(SyspendCrypt);
   i := 1;

    while i < LengthSyspendCrypt do    //������ ������� ����� ����� ������������ ������ �����
       begin
         Syspend := SyspendCrypt[i];
         SyspendCrypt[i] := SyspendCrypt[i+1];
         SyspendCrypt[i+1] := Syspend;
         i := i + 2;
       end;

      while i < LengthSyspendCrypt do  //������ ������� � ���� ��������� ���� � ������
       begin
         Syspend := SyspendCrypt[i];
         SyspendCrypt[i] := SyspendCrypt[Round(LengthSyspendCrypt/2)+i];
         SyspendCrypt[Round(LengthSyspendCrypt/2)+i] := Syspend;
         i := i + 2;
       end;

     Result := SyspendCrypt;

   end;

 //*********************** �������� ����������� ��������� *******************************

   procedure CheckCrypt;


   begin






   end;





 //*********************** MAIN PROGRAMM *******************************************************

begin
   ProgramName:='Blevantone 5.0 local crypt';
   ProgBar:=0;


  //    Form1.Label2.Caption:=Label2C;
  // TForm1.Caption:= ProgramName;
  //   Form1.Caption:= ProgramName;
  // IniFile:=TIniFile.Create(ExtractFilePath(Application.ExeName)+ProgramName+'.ini');
  // CheckCrypt.Form2.Show;
  //ShowMessage(ExtractFilePath(Application.ExeName));
   //ShowMessage(GetHardDiskSerial('c'));
   ShowMessage(Password+'password');
   // Form2.ShowModal;


  // Form2.Show; // (��� Form2.ShowModal) ����� �����



    ShowMessage('WordActivate');
  IniFileCreate;








end.

