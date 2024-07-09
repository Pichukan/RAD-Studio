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


//******** ��������� �������� ��������� EXCEL �� ��������� ���� **********************

procedure ActExcelOpenDoc;
begin

  DIRExName:=DIRName+'xlsm';

  MyBook:= EXC.WorkBooks.Open(DIRExName);  //��������� ��������
//  ShowMessage(DIRExName);   //���������� ���������� ��������� ����� � ������ � �����������

end;



//******** ��������� ����������� ���� � ��������� EXCEL    *****************************
//********         �� ������� ����������� ����������       ******************************

procedure DetectExcelOpenExpertDoc;

 var
 CountDir  : Integer;

begin

 // ShowMessage(DIRFileDetect);

  CountDir := Length(DIRFileDetect);

  Delete(DIRName, CountDir+1, 2);

  DIRExName:=DIRName+'xlsm';

 // ShowMessage(DIRExName);   //���������� ���������� ��������� ����� � ������ � �����������

end;


//**********              ��� �������, ������, ��, ����������                  ***********
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


//**********              ��� ������ � ��������� ��������                     ***********
//********** ��������� ��������� ������ EXCEL ���������� ������� ��� �������  ***********
//********** �� ����� EXCEL � ����������� �� ��� ��, � ����� ��������� ���    ***********
//********** ������� ��������� � �������� WORD �� �������� ����� ����� ������ ***********
//********** ��������� ������                                                 ***********

procedure TableAsPicturePasteBuildLand(Book : OleVariant);
  var  ReplaceText : array [1..30] of string;
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
     ReplaceText [1]:='1#1��';        //���������� �� ������� ����� � WORD
     ReplaceText [2]:='1#2��';        //���������� �� ������� ����� � WORD
     ReplaceText [3]:='1#3��';        //���������� �� ������� ����� � WORD
     ReplaceText [4]:='1#4��';        //���������� �� ������� ����� � WORD
     ReplaceText [5]:='1#5��';        //���������� �� ������� ����� � WORD
     ReplaceText [6]:='1#6��';        //���������� �� ������� ����� � WORD

     ReplaceText [7]:='1#1����';        //���������� �� ������� ����� � WORD
     ReplaceText [8]:='1#2����';        //���������� �� ������� ����� � WORD
     ReplaceText [9]:='1#3����';        //���������� �� ������� ����� � WORD
     ReplaceText [10]:='1#4����';        //���������� �� ������� ����� � WORD
     ReplaceText [11]:='1#5����';        //���������� �� ������� ����� � WORD
     ReplaceText [12]:='1#6����';        //���������� �� ������� ����� � WORD

  //   ReplaceText [13]:='1#1����';      //���������� �� ������� ����� � WORD
  //   ReplaceText [14]:='1#2����';      //���������� �� ������� ����� � WORD
     ReplaceText [13]:='1#3����';        //���������� �� ������� ����� � WORD
     ReplaceText [14]:='1#4����';        //���������� �� ������� ����� � WORD
     ReplaceText [15]:='1#5����';        //���������� �� ������� ����� � WORD
     ReplaceText [16]:='1#6����';        //���������� �� ������� ����� � WORD
     ReplaceText [17]:='1#7����';        //���������� �� ������� ����� � WORD
     ReplaceText [18]:='1#8����';        //���������� �� ������� ����� � WORD
     ReplaceText [19]:='1#9����';        //���������� �� ������� ����� � WORD
     ReplaceText [20]:='1#10����';       //���������� �� ������� ����� � WORD

     ReplaceText [21]:='1#1����';        //���������� �� ������� ����� � WORD
     ReplaceText [22]:='1#2����';        //���������� �� ������� ����� � WORD
     ReplaceText [23]:='1#3����';        //���������� �� ������� ����� � WORD
     ReplaceText [24]:='1#4����';        //���������� �� ������� ����� � WORD

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


 //***************������� �������� ��� ������� � ������� �� EXCEL � ������*********************************************************************


 //******* ��������� ���������  ������� ��� ������� � ������� �� ��������� EXCEL   ********
//******* � �������� � WORD   ��� ��������� ������            ********
//******* (          ********

 procedure InsertPictureWordZad;
  var
    j,i : Integer;
 begin

     BookZad:=WRD.Documents.Open(DIRFile+'������.docx');

     MyWorkSheet:=MyBook.Sheets['�������'];
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




//************������� �� ���� � ����� WORD*****************************************


//******* ��������� ���������  ������� �� ��������� WORD   ********
//******* � ������ �������� � �������� �����               ********
//******* (� ������ ������ ��� ���. � ��� �������          ********
//procedure InsertPictureWord(Book : OleVariant);
 procedure InsertPictureWord();
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
             ShowMessage('����� � ���������;'+ReplaceTextValue+' �� ������.');
             BookPicture.Close;
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

       BookPicture.Close


end;



//********************��������� ����� �� �������� � EXCEL � ����� WORD**************************************************************************

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
        DirObzor:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [1,2],'�������� �� ����� ����������')+ObzorValue+' �����.docx';

     //ShowMessage(DirObzor);

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



 //*********************��������� ����� ������� �� �������� � EXCEL � ����� WORD************************************

//**** ��������� ���� �� ������� EXCEL �������� ������� ������  obj             *****
//**** ��������� �������� WORD � ������� �������� �� �������� ���������       *****
//**** � ��������� � �������� WORD ������                            *****
 procedure InsertObzorObj();
 var
    DirObzorObj,ReplaceTextObzorObj : string;
    DefaultReadIniFileObj        : string;

 begin
      MyWorkSheet2:=MyBook.Sheets['����'];
      RangeObzorObj:=MyWorkSheet2.Range['c3'];
      ObzorValueObj:= RangeObzorObj.Value;

    //ShowMessage(vartostr(ObzorValue));

        DefaultReadIniFileObj := 'Z:\GRAND NEVA\2014\������ ������Z\';   //�������� ���������� ������ �� ���������
        DirObzorObj:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [4,2], '���������� ��������')+ObzorValueObj+'.docx';

     //ShowMessage(DirObzorObj);

      try
      BookObzorObj:=WRD.Documents.Open(DirObzorObj);
      except

         ShowMessage('BLEVANTONE �� ���� ������� ���� � ������� �������.'+
       ' ���������� ������� ������������ ��������� � ������� ������� '+
        ObzorValueObj+' �� ���, ���� ���, ������ ������� ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //��� ����������� ������ ������ �����, � ���� ������������ ����� "Cancel", �� �������

          DirObzorObj := Form1.OpenDialog1.FileName;
          BookObzorObj:=WRD.Documents.Open(DirObzorObj);

         //******* ����������� � ������ � INI ���� ���������� ����� � �������� *****

         IniSectionValue [4,3]:= DIRFileDetect2(DirObzorObj);  //��������� ����� ����� � ������� ������ �� ����
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [4,2], IniSectionValue [4,3]);

      end;

      BookObzorObj.Range.Copy;
      ReplaceTextObzorObj:='7###obj';
      MyRange2 := FindInDoc(Book, ReplaceTextObzorObj);

         if VarIsClear(MyRange2) then begin
            ShowMessage('����� 7###obj ReplaceTextObzorObj �� ������.');
            Exit;
         end;

  MyRange2.Paste;
 end;





//********************��������� ����� �� � ������ � ����� WORD******************************************************


//**** ��������� ���� �� ������� EXCEL �������� ������� ������  obj             *****
//**** ��������� �������� WORD � ������� �������� �� �������� ���������       *****
//**** � ��������� � �������� WORD ������                            *****
 procedure InsertObzorRFandRegion();
 var
    DirObzorRF, DirObzorRegion, ReplaceTextObzorRF : string;
    ReplaceTextObzorRegion : string;
    DefaultReadIniFileRF,DefaultReadIniFileRegion         : string;

 begin
     // MyWorkSheet2:=MyBook.Sheets['����'];
     // RangeObzorObj:=MyWorkSheet2.Range['c3'];
      ObzorValueRF:= '����� ��';
      ObzorValueRegion:= '����� ������';

    //ShowMessage(vartostr(ObzorValue));

        DefaultReadIniFileRF := 'Z:\GRAND NEVA\2014\������ ������Z\';   //�������� ���������� ������ �� ���������
        DirObzorRF:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [5,2], '���������� ��������')+ObzorValueRF+'.docx';


        DefaultReadIniFileRegion := 'Z:\GRAND NEVA\2014\������ ������Z\';   //�������� ���������� ������ �� ���������
        DirObzorRegion:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [6,2], '���������� ��������')+ObzorValueRegion+'.docx';

     //ShowMessage(DirObzorObj);

      try
      BookObzorRF:=WRD.Documents.Open(DirObzorRF);
      except

         ShowMessage('BLEVANTONE �� ���� ������� ���� � ������� RF.'+
       ' ���������� ������� ������������ ��������� � ������� RF '+
        ObzorValueRF+' �� ���, ���� ���, ������ ������� ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //��� ����������� ������ ������ �����, � ���� ������������ ����� "Cancel", �� �������

          DirObzorRF := Form1.OpenDialog1.FileName;
          BookObzorRF:=WRD.Documents.Open(DirObzorRF);

         //******* ����������� � ������ � INI ���� ���������� ����� � �������� *****

         IniSectionValue [5,3]:= DIRFileDetect2(DirObzorRF);  //��������� ����� ����� � ������� ������ �� ����
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [5,2], IniSectionValue [5,3]);

      end;



      try
      BookObzorRegion:=WRD.Documents.Open(DirObzorRegion);
      except

         ShowMessage('BLEVANTONE �� ���� ������� ���� � ������� Region.'+
       ' ���������� ������� ������������ ��������� � ������� Region '+
        ObzorValueRegion+' �� ���, ���� ���, ������ ������� ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //��� ����������� ������ ������ �����, � ���� ������������ ����� "Cancel", �� �������

          DirObzorRegion := Form1.OpenDialog1.FileName;
          BookObzorRegion:=WRD.Documents.Open(DirObzorRegion);

         //******* ����������� � ������ � INI ���� ���������� ����� � �������� *****

         IniSectionValue [6,3]:= DIRFileDetect2(DirObzorRegion);  //��������� ����� ����� � ������� ������ �� ����
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [6,2], IniSectionValue [6,3]);

      end;


      BookObzorRF.Range.Copy;
      ReplaceTextObzorRF:='8###rf';
      MyRange2 := FindInDoc(Book, ReplaceTextObzorRF);

         if VarIsClear(MyRange2) then begin
            ShowMessage('����� 8###rf ReplaceTextObzorRF �� ������.');
            Exit;
         end;

  MyRange2.Paste;


  BookObzorRegion.Range.Copy;
      ReplaceTextObzorRegion:='9###reg';
      MyRange2 := FindInDoc(Book, ReplaceTextObzorRegion);

         if VarIsClear(MyRange2) then begin
            ShowMessage('����� 9###reg ReplaceTextObzorRegion �� ������.');
            Exit;
         end;

  MyRange2.Paste;
 end;



//*****************��������� ��������� ����� ��������� � ����� WORD******************************************


//**** ��������� ���� �� ����� ini ���� � ����� ��� ��� ��� (���������, ������������ � ��. *****
//**** ��������� �������� WORD � ����������� (�����, ���� � ��) �� �������� ���������       *****
//**** � ��������� � �������� WORD ������                            *****
 procedure InsertDociOcen();
 var
    DirDoci,ReplaceTextDoci : string;
    DefaultReadIniFileDoci        : string;

 begin
     // MyWorkSheet2:=MyBook.Sheets['����'];
     // RangeObzorObj:=MyWorkSheet2.Range['c3'];
      ObzorValueDoci:= '��� �������';

    //ShowMessage(vartostr(ObzorValue));

        DefaultReadIniFileDoci := 'Z:\GRAND NEVA\2014\������ ������Z\';   //�������� ���������� ������ �� ���������
        DirDoci:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [7,2], '���������� ��������');

     //ShowMessage(DirObzorObj);

      try
      BookObzorDoci:=WRD.Documents.Open(DirDoci);
      except

         ShowMessage('BLEVANTONE �� ���� ������� ���� � �����������.'+
       ' ���������� ������� ������������ ��������� � ����������� '+
        ObzorValueDoci+' �� ���, ���� ���, ������ ������� ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //��� ����������� ������ ������ �����, � ���� ������������ ����� "Cancel", �� �������

          DirDoci := Form1.OpenDialog1.FileName;
          BookObzorDoci:=WRD.Documents.Open(DirDoci);

         //******* ����������� � ������ � INI ���� ���������� ����� � �������� *****

         IniSectionValue [7,3]:= DirDoci;  //��� ��������� ����� ����� � ������� ������ �� ����
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [7,2], IniSectionValue [7,3]);

      end;

      BookObzorDoci.Range.Copy;
      ReplaceTextDoci:='10###doc';
      MyRange2 := FindInDoc(Book, ReplaceTextDoci);

         if VarIsClear(MyRange2) then begin
            ShowMessage('����� 10###doc ReplaceTextDoci �� ������.');
            Exit;
         end;

  MyRange2.Paste;
 end;






 //****************��������� � ����� ��������� ���������**********************************************************

//**** ��������� ���� �� ����� ini ���� � ����� ��� ��� ��� (���������, ������������ � ��. *****
//**** ��������� �������� WORD � ����������� (�����, ���� � ��) �� �������� ���������       *****
//**** � ��������� � �������� WORD ������                            *****
 procedure InsertDociRec();
 var
    DirRec,ReplaceTextRec : string;
    DefaultReadIniFileRec        : string;

 begin
     // MyWorkSheet2:=MyBook.Sheets['����'];
     // RangeObzorObj:=MyWorkSheet2.Range['c3'];
      ObzorValueRec:= '��� ������� ���������';

    //ShowMessage(vartostr(ObzorValue));

        DefaultReadIniFileRec := 'Z:\GRAND NEVA\2014\������ ������Z\';   //�������� ���������� ������ �� ���������
        DirRec:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [8,2], '���������� ��������');

     //ShowMessage(DirObzorObj);

      try
      BookObzorRec:=WRD.Documents.Open(DirRec);
      except

         ShowMessage('BLEVANTONE �� ���� ������� ���� � ����������� ��������.'+
       ' ���������� ������� ������������ ��������� � ����������� '+
        ObzorValueRec+' �� ���, ���� ���, ������ ������� ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //��� ����������� ������ ������ �����, � ���� ������������ ����� "Cancel", �� �������

          DirRec := Form1.OpenDialog1.FileName;
          BookObzorRec:=WRD.Documents.Open(DirRec);

         //******* ����������� � ������ � INI ���� ���������� ����� � �������� *****

         IniSectionValue [8,3]:= DirRec;  //��� ��������� ����� ����� � ������� ������ �� ����
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [8,2], IniSectionValue [8,3]);

      end;

      BookObzorRec.Range.Copy;
      ReplaceTextRec:='11###rec';
      MyRange2 := FindInDoc(Book, ReplaceTextRec);

         if VarIsClear(MyRange2) then begin
            ShowMessage('����� 11###rec � ReplaceTextRec �� ������.');
            Exit;
         end;

  MyRange2.Paste;
 end;




 //****************��������� ��� � ������������ � ����� WORD********************************************************************************************************


//**** ��������� ���� �� ����� ini ���� � ��������� � ������������ ����������� ������ (����������������) *****
//**** ��������� �������� WORD � ����������������� �� �������� ���������       *****
//**** � ��������� � �������� WORD ������                            *****
 procedure InsertFco();
 var
    DirFco,ReplaceTextFco : string;
    DefaultReadIniFileFco        : string;

 begin
     // MyWorkSheet2:=MyBook.Sheets['����'];
     // RangeObzorObj:=MyWorkSheet2.Range['c3'];
      ObzorValueFco:= '���';

    //ShowMessage(vartostr(ObzorValue));

        DefaultReadIniFileFco := 'Z:\GRAND NEVA\2014\������ ������Z\';   //�������� ���������� ������ �� ���������
        DirFco:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [9,2], '���������� ��������');

     //ShowMessage(DirObzorObj);

      try
      BookObzorFco:=WRD.Documents.Open(DirFco);
      except

         ShowMessage('BLEVANTONE �� ���� ������� ���� � ���.'+
       ' ���������� ������� ������������ ��������� � ��� '+
        ObzorValueFco+' �� ���, ���� ���, ������ ������� ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //��� ����������� ������ ������ �����, � ���� ������������ ����� "Cancel", �� �������

          DirFco := Form1.OpenDialog1.FileName;
          BookObzorFco:=WRD.Documents.Open(DirFco);

         //******* ����������� � ������ � INI ���� ���������� ����� � �������� *****

         IniSectionValue [9,3]:= DirFco;  //��� ��������� ����� ����� � ������� ������ �� ����
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [9,2], IniSectionValue [9,3]);

      end;

      BookObzorFco.Range.Copy;

     // ShowMessage('��� � ������ �������');

      ReplaceTextFco:='���###';
      MyRange2 := FindInDoc(Book, ReplaceTextFco);

         if VarIsClear(MyRange2) then begin
            ShowMessage('����� ���### � ReplaceTextFco �� ������.');
            Exit;
         end;

  MyRange2.Paste;

  //ShowMessage('MyRange2 ��� ��������');

 end;





//********************��������� ������ ���������� � ����� WORD*********************************************************************************************


//**** ��������� ���� �� ����� ini ���� � ��������� �� ������� ������������ ���������� *****
//**** ��������� �������� WORD � ������� ���������� �� �������� ���������              *****
//**** � ��������� � �������� WORD ������ �� �������� ����� �� �����                   *****
 procedure InsertLit();
 var
    DirLit,ReplaceTextLit : string;
    DefaultReadIniFileLit        : string;

 begin
     // MyWorkSheet2:=MyBook.Sheets['����'];
     // RangeObzorObj:=MyWorkSheet2.Range['c3'];
      ObzorValueLit:= '����������';

    //ShowMessage(vartostr(ObzorValue));

        DefaultReadIniFileLit := 'Z:\GRAND NEVA\2014\����������\';   //�������� ���������� ������ �� ���������
        DirLit:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [13,2], '���������� ��������');

     //ShowMessage(DirObzorObj);

      try
      BookObzorLit:=WRD.Documents.Open(DirLit);
      except

         ShowMessage('BLEVANTONE �� ���� ������� ���� � �����������.'+
       ' ���������� ������� ������������ ��������� � ����������� '+
        ObzorValueLit+' �� ���, ���� ���, ������ ������� ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //��� ����������� ������ ������ �����, � ���� ������������ ����� "Cancel", �� �������

          DirLit := Form1.OpenDialog1.FileName;
          BookObzorLit:=WRD.Documents.Open(DirLit);

         //******* ����������� � ������ � INI ���� ���������� ����� � �������� *****

         IniSectionValue [13,3]:= DirLit;  //��� ��������� ����� ����� � ������� ������ �� ����
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [13,2], IniSectionValue [13,3]);

      end;

      BookObzorLit.Range.Copy;

     // ShowMessage('���������� � ������ �������');

      ReplaceTextLit:='�����###';
      MyRange2 := FindInDoc(Book, ReplaceTextLit);

         if VarIsClear(MyRange2) then begin
            ShowMessage('����� �����### � ReplaceTextLit �� ������.');
            Exit;
         end;

  MyRange2.Paste;

  //ShowMessage('MyRange2 lit ��������');

 end;




//********************��������� ����������� � ������������� � ����� WORD*******************************************************************************************


//**** ��������� ���� �� ����� ini ���� � ������ �� ������������ � �������������  *****
//**** ��������� �������� WORD � ������������� � ��������������� �� �������� ���������       *****
//**** � ��������� � �������� WORD ������                            *****
 procedure InsertOgr();
 var
    DirOgr,ReplaceTextOgr : string;
    DefaultReadIniFileOgr        : string;

 begin
     // MyWorkSheet2:=MyBook.Sheets['����'];
     // RangeObzorObj:=MyWorkSheet2.Range['c3'];
      ObzorValueOgr:= '����������� � �������������';

    //ShowMessage(vartostr(ObzorValue));

        DefaultReadIniFileOgr := 'Z:\GRAND NEVA\2014\������ ������Z\';   //�������� ���������� ������ �� ���������
        DirOgr:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [10,2], '���������� ��������');

     //ShowMessage(DirObzorObj);

      try
      BookObzorOgr:=WRD.Documents.Open(DirOgr);
      except

         ShowMessage('BLEVANTONE �� ���� ������� ���� � �������������.'+
       ' ���������� ������� ������������ ��������� � ������������� '+
        ObzorValueOgr+' �� ���, ���� ���, ������ ������� ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //��� ����������� ������ ������ �����, � ���� ������������ ����� "Cancel", �� �������

          DirOgr := Form1.OpenDialog1.FileName;
          BookObzorOgr:=WRD.Documents.Open(DirOgr);

         //******* ����������� � ������ � INI ���� ���������� ����� � �������� *****

         IniSectionValue [10,3]:= DirOgr;  //��� ��������� ����� ����� � ������� ������ �� ����
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [10,2], IniSectionValue [10,3]);

      end;

      BookObzorOgr.Range.Copy;
      ReplaceTextOgr:='�����������###';
      MyRange2 := FindInDoc(Book, ReplaceTextOgr);

         if VarIsClear(MyRange2) then begin
            ShowMessage('����� �����������### � ReplaceTextOgr �� ������.');
            Exit;
         end;

  MyRange2.Paste;
 end;




//***************��������� ��������� � ����� WORD***********************************************************


//**** ��������� ���� �� ����� ini ���� � ������ �� ����������   *****
//**** ��������� �������� WORD � ����������� �� �������� ���������       *****
//**** � ��������� � �������� WORD ������                            *****
 procedure InsertDop();
 var
    DirDop,ReplaceTextDop : string;
    DefaultReadIniFileDop        : string;

 begin
     // MyWorkSheet2:=MyBook.Sheets['����'];
     // RangeObzorObj:=MyWorkSheet2.Range['c3'];
      ObzorValueDop:= '������ ���������';

    //ShowMessage(vartostr(ObzorValue));

        DefaultReadIniFileDop := 'Z:\GRAND NEVA\2014\������ ������Z\';   //�������� ���������� ������ �� ���������
        DirDop:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [11,2], '���������� ��������');

     //ShowMessage(DirObzorObj);

      try
      BookObzorDop:=WRD.Documents.Open(DirDop);
      except

         ShowMessage('BLEVANTONE �� ���� ������� ���� � �����������.'+
       ' ���������� ������� ������������ ��������� � ����������� '+
        ObzorValueDop+' �� ���, ���� ���, ������ ������� ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //��� ����������� ������ ������ �����, � ���� ������������ ����� "Cancel", �� �������

          DirDop := Form1.OpenDialog1.FileName;
          BookObzorDop:=WRD.Documents.Open(DirDop);

         //******* ����������� � ������ � INI ���� ���������� ����� � �������� *****

         IniSectionValue [11,3]:= DirDop;  //��� ��������� ����� ����� � ������� ������ �� ����
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [11,2], IniSectionValue [11,3]);

      end;

      BookObzorDop.Range.Copy;
      ReplaceTextDop:='���������###';
      MyRange2 := FindInDoc(Book, ReplaceTextDop);

         if VarIsClear(MyRange2) then begin
            ShowMessage('����� ���������### � ReplaceTextDop �� ������.');
            Exit;
         end;

  MyRange2.Paste;
 end;





//**************��������� ������� � ����������� � ����� WORD*******************************************


//**** ��������� ���� �� ����� ini ���� � ������ �� ������� � �����������   *****
//**** ��������� �������� WORD � ��������� �� �������� ���������       *****
//**** � ��������� � �������� WORD ������                            *****
 procedure InsertTerm();
 var
    DirTerm,ReplaceTextTerm : string;
    DefaultReadIniFileTerm        : string;

 begin
     // MyWorkSheet2:=MyBook.Sheets['����'];
     // RangeObzorObj:=MyWorkSheet2.Range['c3'];
      ObzorValueTerm:= '������� � �����������';

    //ShowMessage(vartostr(ObzorValue));

        DefaultReadIniFileTerm := 'Z:\GRAND NEVA\2014\������ ������Z\';   //�������� ���������� ������ �� ���������
        DirTerm:= IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [12,2], '���������� ��������');

     //ShowMessage(DirObzorObj);

      try
      BookObzorTerm:=WRD.Documents.Open(DirTerm);
      except

         ShowMessage('BLEVANTONE �� ���� ������� ���� � ��������� � �������������.'+
       ' ���������� ������� ������������ ��������� � ��������� � ������������� '+
        ObzorValueTerm+' �� ���, ���� ���, ������ ������� ;)');

        if not Form1.OpenDialog1.Execute then Exit;  //��� ����������� ������ ������ �����, � ���� ������������ ����� "Cancel", �� �������

          DirTerm := Form1.OpenDialog1.FileName;
          BookObzorTerm:=WRD.Documents.Open(DirTerm);

         //******* ����������� � ������ � INI ���� ���������� ����� � �������� *****

         IniSectionValue [12,3]:= DirTerm;  //��� ��������� ����� ����� � ������� ������ �� ����
         IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [12,2], IniSectionValue [12,3]);

      end;

      BookObzorTerm.Range.Copy;
      ReplaceTextTerm:='�������###';
      MyRange2 := FindInDoc(Book, ReplaceTextTerm);

         if VarIsClear(MyRange2) then begin
            ShowMessage('����� �������### � ReplaceTextTerm �� ������.');
            Exit;
         end;

  MyRange2.Paste;
 end;




 //****************��������� ������� � ����� WORD********************************************************************************************************


  //***       ������� ��� �������, ������, ��, ����������         ***
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


 //***       ������� ��� ������ � ��������� ��������              ***
 //**** ��������� ��������� ������� �������� �� ��������� WORD    ****
 //****                 � �������� ������                         ****
 procedure InsertAnalogiBuildLand;
  var
        DirAnalogDoc, ReplaceTextAnalog : string;

 begin

    DirAnalogDoc:= DIRFile+'������� ��.docx';

    try
      BookAnalog:=WRD.Documents.Open(DirAnalogDoc);
      except
      ShowMessage('�� ���������� ������� ���� � ��������� ��.'+
      '��������� �������, ������������ �����');
      Exit;
    end;

    BookAnalog.Range.Copy;
    ReplaceTextAnalog:='1######1';
    MyRange2 := FindInDoc(Book, ReplaceTextAnalog);

         if VarIsClear(MyRange2) then begin
            ShowMessage('����� 1######1 ReplaceTextAnalog �� ������.');
            Exit;
         end;

    MyRange2.Paste;                //������� ��

    DirAnalogDoc:= DIRFile+'������� ���+��.docx';

    try
      BookAnalogBuild:=WRD.Documents.Open(DirAnalogDoc);
      except
      ShowMessage('�� ���������� ������� ���� � ��������� ���+��.'+
      '��������� �������, ������������ �����');
      Exit;
    end;

    BookAnalogBuild.Range.Copy;
    ReplaceTextAnalog:='1######2';
    MyRange2 := FindInDoc(Book, ReplaceTextAnalog);

         if VarIsClear(MyRange2) then begin
            ShowMessage('����� 1######2 ReplaceTextAnalog �� ������.');
            Exit;
         end;

    MyRange2.Paste;                 //������� ���+��
    BookAnalogBuild.Close;

 end;



 //**************��������� ������ ������� � ����� WORD****************************************************

//*******  ������� ��� ������ �� �� ������ �����������              *******
//******* ��������� ���������  ������� ������ �� ������             ********
//******* �� ��������� WORD � ������ �������� � �������� �����      ********
//******* (� ������ ������ ��� � ������ ����� � ��������� �� ������ ********

procedure InsertAnalogiCoInvest;

  var
        DirAnalogDoc, ReplaceTextAnalog : string;
        j                               : Integer;

 begin

    DirAnalogDoc:= DIRFile+'������� ��-������.docx';

    try
      BookAnalogiCoInvest:=WRD.Documents.Open(DirAnalogDoc);
      except
      ShowMessage('�� ���������� ������� ���� � ��������� �� ������.'+
      '��������� �������, ������������ �����');
      Exit;
    end;
   //
    j:=BookAnalogiCoInvest.InlineShapes.Count;

     case j of
      0 : ShowMessage('� �������� �� ������ ������ ���');
     end;

    if j=0  then   Exit;
   //
    BookAnalogiCoInvest.Range.Copy;

    ReplaceTextAnalog:='1#1����';

    MyRange2 := FindInDoc(Book, ReplaceTextAnalog);

         if VarIsClear(MyRange2) then begin
            ShowMessage('����� 1#1���� ReplaceTextAnalog �� ������.');
            Exit;
         end;

    MyRange2.Paste;

    BookAnalogiCoInvest.Close;

end;



//*****************��������� ���� ������� ������ � ����� WORD*******************************************


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



//**************��������� ������� �������������� � ����� WORD**********************************************


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
            ShowMessage('����� 1####� Location ���������� �� ���������� �������.');
            Exit;
         end;

        // Sleep(1000);
         MyRange2.Paste;
       end;

  end;


 //**** ������ ������� �� ������� ��������� � �������� EXCEL    ************************************

   procedure MacroSli;

   begin
      Book.Activate;
      WRD.Run('�������');

   end;


    procedure MacroSliExpert;

   begin
      Book.Activate;
      WRD.Run('��������������');

   end;

   procedure MacroSliDoc;

   begin
      Book.Activate;
      WRD.Run('�����������������');

   end;


   procedure MacroSliCalc;

   begin
      Book.Activate;
      WRD.Run('��������������');

   end;


   procedure ExeMacros;

    begin
     // MyBook.Activate;
      EXC.Run('����������');

    end;

   procedure ExeMacrosZad;

    begin
     // MyBook.Activate;
      EXC.Run('������');

    end;

 //**** ��������� ��������� ��������� WORD ����� ������**************************************************
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
    //BookPicture.Close;    ��������� � ����� ������������, ��� ���
    //                      ���� ��������� ����� �� �������� ��������
    //                      ������, �� ���� ����������� ������ ��, ���
    //                      ������ ������ � �������� ������� ������
    //                      �� ���� ������ �������. ���� ��� ��������� �
    //                      ���������� ������� �� ����� ����� ���.
   end;

 //**** ��������� ��������� �������� EXCEL
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
ShowMessage('����� ����� ������� �����... � ����� �� �������� �����������.');
end;

procedure TForm1.WhatDoClick(Sender: TObject);
begin
ShowMessage(' �������� ����� ������� ������');
end;


//********** ������� ������ Start Program (��������, �������, ��, ��������� *********

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

    //TableAsPicturePaste(Book);
    TableAsPicturePaste(Book);

     ProgBar:=33;
    ProgressBar1.Position := ProgBar ;
    Label2C:='������� EXCEL ���������, ������� ���������';
    Form1.Label2.Caption:=Label2C;


    //InsertPictureWord(Book);
      InsertPictureWord();

     ProgBar:=40;
    ProgressBar1.Position := ProgBar ;
    Label2C:='���, ����, ������� ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

    //BookPicture.Close;

    InsertObzor() ;

      ProgBar:=45;
    ProgressBar1.Position := ProgBar ;
    Label2C:='����� ������ ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

     InsertObzorObj() ;

      ProgBar:=51;
    ProgressBar1.Position := ProgBar ;
    Label2C:='����� obj ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

     InsertObzorRFandRegion();

     ProgBar:=56;
    ProgressBar1.Position := ProgBar ;
    Label2C:='����� RF � Region ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

     InsertFco();

     ProgBar:=60;
    ProgressBar1.Position := ProgBar ;
    Label2C:='����� ��� ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

    BookObzorFco.Close();


   InsertOgr();

     ProgBar:=63;
    ProgressBar1.Position := ProgBar ;
    Label2C:='����������� � ������������� ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

    BookObzorOgr.Close();


    InsertDop();

     ProgBar:=66;
    ProgressBar1.Position := ProgBar ;
    Label2C:='��������� ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

    BookObzorDop.Close();


    InsertTerm();

     ProgBar:=68;
    ProgressBar1.Position := ProgBar ;
    Label2C:='������� � ����������� ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

    BookObzorTerm.Close();


    InsertLit();

     ProgBar:=69;
    ProgressBar1.Position := ProgBar ;
    Label2C:='���������� ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

    BookObzorLit.Close();




    InsertDociOcen();

    ProgBar:=71;
    ProgressBar1.Position := ProgBar ;
    Label2C:='���� ��������� ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

     InsertDociRec();

    ProgBar:=73;
    ProgressBar1.Position := ProgBar ;
    Label2C:='��������� ��������� ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

     CloseExcel;

     ProgBar:=75;
    ProgressBar1.Position := ProgBar ;
    Label2C:='������ EXCEL, ������� ��������� �������';
    Form1.Label2.Caption:=Label2C;

    InsertAnalogi;

      ProgBar:=79;
    ProgressBar1.Position := ProgBar ;
    Label2C:='������� ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

    InsertFoto;

      ProgBar:=82;
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

//********* ������� ������ ���������� ���������� ************************

   procedure TForm1.Button3Click(Sender: TObject);
begin

     ProgBar:=0;
     ProgressBar1.Position := ProgBar ;
     Label2C:='��������� ����������� ���������� ������ ������';
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

      ProgBar:=25;
    ProgressBar1.Position := ProgBar ;
     Label2C:='�������� WORD �� ��������� ���������� ������';
     Form1.Label2.Caption:=Label2C;

  //  StartExcel;

 //   Label2C:='���������� EXCEL ����������';
 //   Form1.Label2.Caption:=Label2C;

    DetectExcelOpenExpertDoc;             //���������� ��� ����� EXCEL �� �������� ����������
                                          //���� EXCEL �� ���������
     ProgBar:=45;
    ProgressBar1.Position := ProgBar ;
     Label2C:='�������� EXCEL �� ��������� ���������� �������';
     Form1.Label2.Caption:=Label2C;

      MacroSliExpert;

     ProgBar:=75;
    ProgressBar1.Position := ProgBar ;
     Label2C:='�������� WORD � EXCEL ���������';
    Form1.Label2.Caption:=Label2C;


     ProgBar:=85;
    ProgressBar1.Position := ProgBar ;
     Label2C:='��� ������� � �����';
    Form1.Label2.Caption:=Label2C;


  WRD.visible:=true;     //������ ���� �������

   ProgBar:=100;
    ProgressBar1.Position := ProgBar ;
    Label2C:='��������� '+ProgramName+' ��������� (���������� �����)';
    Form1.Label2.Caption:=Label2C;

  ShowMessage('program '+ ProgramName + ' end');


     ProgBar:=0;
    ProgressBar1.Position := ProgBar ;
     Label2C:='��������� ������� �������';
    Form1.Label2.Caption:=Label2C;



end;

//********* ������� ������ ��� ������� � ������� �� ������ *******************

procedure TForm1.Button4Click(Sender: TObject);
begin
   ProgBar:=0;
     ProgressBar1.Position := ProgBar ;
     Label2C:='��������� ����������� ���� � ������� ������ ������';
     Form1.Label2.Caption:=Label2C;



     StartExcel;

     if not OpenDialog1.Execute then Exit;
     DIRFName := OpenDialog1.FileName;    //���������� ���������� ��������� �����
  //  ShowMessage(DIRFName);             // � ������ � �����������
     MyBook:= EXC.WorkBooks.Open(DIRFName);


    StartWORD;

    ProgBar:=15;
    ProgressBar1.Position := ProgBar ;
    Label2C:='���������� WORD � ������ ����������';
    Form1.Label2.Caption:=Label2C;


  //  ShowMessage(DIRFileDetect);    // ������������ ���������� ����� ��� �������� �����
    DIRFile:= DIRFileDetect;       //� ������������� ����������

      ProgBar:=40;
    ProgressBar1.Position := ProgBar ;
     Label2C:='�������� WORD �� ��������� ���������� ������';
     Form1.Label2.Caption:=Label2C;

   ExeMacrosZad;

   InsertPictureWordZad;

   CloseExcel;

  WRD.visible:=true;     //������ ���� �������

   ProgBar:=100;
    ProgressBar1.Position := ProgBar ;
    Label2C:='��������� '+ProgramName+' ��������� (������� � ��� ���������)';
    Form1.Label2.Caption:=Label2C;

  ShowMessage('program '+ ProgramName + ' end');

     ProgBar:=0;
    ProgressBar1.Position := ProgBar ;
     Label2C:='��������� ������� �������';
    Form1.Label2.Caption:=Label2C;


end;




//********* ������� ������ ������� ��������� *******************

procedure TForm1.Button5Click(Sender: TObject);
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

 {   ExeMacros;

    Label2C:='������ EXCEL ���� ��� ��� ��������';
    Form1.Label2.Caption:=Label2C;
    EXC.visible:=true;       //������ Excel �������
 // EXC.visible:=false;      //������ Excel ���������

 }
      ProgBar:=17;
    ProgressBar1.Position := ProgBar ;

 {   TableAsPicturePasteBuildLand(Book);    //*******************************
  }
    ProgBar:=33;
    ProgressBar1.Position := ProgBar ;
  {  Label2C:='������� EXCEL ���������, ������� ���������';
    Form1.Label2.Caption:=Label2C;
  }

    //InsertPictureWord(Book);
    InsertPictureWord();

     ProgBar:=45;
    ProgressBar1.Position := ProgBar ;
    Label2C:='���, ����, ������� ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

    InsertObzor() ;

      ProgBar:=50;
    ProgressBar1.Position := ProgBar ;
    Label2C:='����� ������ ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

     InsertObzorObj() ;

      ProgBar:=58;
    ProgressBar1.Position := ProgBar ;
    Label2C:='����� obj ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

    InsertObzorRFandRegion();

     ProgBar:=62;
    ProgressBar1.Position := ProgBar ;
    Label2C:='����� RF � Region ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

    InsertDociOcen();

    ProgBar:=69;
    ProgressBar1.Position := ProgBar ;
    Label2C:='���� ��������� ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

     InsertDociRec();

    ProgBar:=71;
    ProgressBar1.Position := ProgBar ;
    Label2C:='��������� ��������� ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

     CloseExcel;

     ProgBar:=73;
    ProgressBar1.Position := ProgBar ;
    Label2C:='������ EXCEL, ������� ��������� �������';
    Form1.Label2.Caption:=Label2C;

  {  InsertAnalogiBuildLand;
  }
   { InsertAnalogiCoInvest;
    }
      ProgBar:=75;
    ProgressBar1.Position := ProgBar ;
 {   Label2C:='������� ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;
  }
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

  //  CloseWordDocs;

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





//********* ������� ������ ������ � ��������� �������� *******************

procedure TForm1.Button1Click(Sender: TObject);
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

    TableAsPicturePasteBuildLand(Book);

     ProgBar:=33;
    ProgressBar1.Position := ProgBar ;
    Label2C:='������� EXCEL ���������, ������� ���������';
    Form1.Label2.Caption:=Label2C;


    //InsertPictureWord(Book);
    InsertPictureWord();

     ProgBar:=48;
    ProgressBar1.Position := ProgBar ;
    Label2C:='���, ����, ������� ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

    InsertObzor() ;

      ProgBar:=55;
    ProgressBar1.Position := ProgBar ;
    Label2C:='����� ������ ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

     InsertObzorObj() ;

      ProgBar:=58;
    ProgressBar1.Position := ProgBar ;
    Label2C:='����� obj ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

    InsertObzorRFandRegion();

     ProgBar:=63;
    ProgressBar1.Position := ProgBar ;
    Label2C:='����� RF � Region ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

    InsertDociOcen();

    ProgBar:=69;
    ProgressBar1.Position := ProgBar ;
    Label2C:='���� ��������� ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

     InsertDociRec();

    ProgBar:=71;
    ProgressBar1.Position := ProgBar ;
    Label2C:='��������� ��������� ��������� � WORD ��������';
    Form1.Label2.Caption:=Label2C;

     CloseExcel;

     ProgBar:=73;
    ProgressBar1.Position := ProgBar ;
    Label2C:='������ EXCEL, ������� ��������� �������';
    Form1.Label2.Caption:=Label2C;

    InsertAnalogiBuildLand;

    InsertAnalogiCoInvest;

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




//********** ��������� ����������� �������� ����� �������
//**********       ���������� ������������ �����

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








//**********������� ������ ������� ������� *****************

procedure TForm1.Button6Click(Sender: TObject);
begin

     ProgBar:=0;
     ProgressBar1.Position := ProgBar ;
     Label2C:='��������� ������� �������� ������ ������';
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


  //  DIRName:= DIRDetect;    // ������������ ���������� � ��������� �����
 //   ShowMessage(DIRName);   // ��� ����������  �� � ������






    //ShowMessage(DIRFileDetect);    // ������������ ���������� ����� ��� �������� �����

    DIRFile:= DIRFolderDetect;       //� ������������� ����������

    //  ShowMessage(DIRFile);
  //    ShowMessage(DIRFname);
      ProgBar:=25;
    ProgressBar1.Position := ProgBar ;
     Label2C:='�������� WORD �� ��������� ���������� ������';
     Form1.Label2.Caption:=Label2C;

    StartExcel;

    Label2C:='���������� EXCEL ����������';
    Form1.Label2.Caption:=Label2C;






    //DetectExcelOpenExpertDoc;             //���������� ��� ����� EXCEL �� �������� ����������
                                          //���� EXCEL �� ���������
     ProgBar:=45;
    ProgressBar1.Position := ProgBar ;
     Label2C:='�������� EXCEL �� ��������� ���������� �������';
     Form1.Label2.Caption:=Label2C;

      MacroSliDoc;

     ProgBar:=75;
    ProgressBar1.Position := ProgBar ;
     Label2C:='�������� WORD � EXCEL ���������';
    Form1.Label2.Caption:=Label2C;


     ProgBar:=85;
    ProgressBar1.Position := ProgBar ;
     Label2C:='��� ������� � �����';
    Form1.Label2.Caption:=Label2C;


  WRD.visible:=true;     //������ ���� �������

   ProgBar:=100;
    ProgressBar1.Position := ProgBar ;
    Label2C:='��������� '+ProgramName+' ��������� (������� �����)';
    Form1.Label2.Caption:=Label2C;

  ShowMessage('program '+ ProgramName + ' end');


     ProgBar:=0;
    ProgressBar1.Position := ProgBar ;
     Label2C:='��������� ������� �������';
    Form1.Label2.Caption:=Label2C;



end;





//**********  ������� ������ ���� ������� *****************************

 procedure TForm1.Button7Click(Sender: TObject);
begin

  ProgBar:=0;
     ProgressBar1.Position := ProgBar ;
     Label2C:='��������� ������� ���� ������ ������';
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


  //  DIRName:= DIRDetect;    // ������������ ���������� � ��������� �����
 //   ShowMessage(DIRName);   // ��� ����������  �� � ������






   // ShowMessage(DIRFileDetect);    // ������������ ���������� ����� ��� �������� �����

    DIRFile:= DIRFolderDetect;       //� ������������� ����������

     // ShowMessage(DIRFile);
  //    ShowMessage(DIRFname);
      ProgBar:=25;
    ProgressBar1.Position := ProgBar ;
     Label2C:='�������� WORD �� ��������� ���������� ������';
     Form1.Label2.Caption:=Label2C;

    StartExcel;

    Label2C:='���������� EXCEL ����������';
    Form1.Label2.Caption:=Label2C;






    //DetectExcelOpenExpertDoc;             //���������� ��� ����� EXCEL �� �������� ����������
                                          //���� EXCEL �� ���������
     ProgBar:=45;
    ProgressBar1.Position := ProgBar ;
     Label2C:='�������� EXCEL �� ��������� ���������� �������';
     Form1.Label2.Caption:=Label2C;

      MacroSliCalc;

     ProgBar:=75;
    ProgressBar1.Position := ProgBar ;
     Label2C:='�������� WORD � EXCEL ���������';
    Form1.Label2.Caption:=Label2C;


     ProgBar:=85;
    ProgressBar1.Position := ProgBar ;
     Label2C:='��� ������� � �����';
    Form1.Label2.Caption:=Label2C;


  WRD.visible:=true;     //������ ���� �������

   ProgBar:=100;
    ProgressBar1.Position := ProgBar ;
    Label2C:='��������� '+ProgramName+' ��������� (���� �����)';
    Form1.Label2.Caption:=Label2C;

  ShowMessage('program '+ ProgramName + ' end');


     ProgBar:=0;
    ProgressBar1.Position := ProgBar ;
     Label2C:='��������� ������� �������';
    Form1.Label2.Caption:=Label2C;





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

   var KeyCrypt       : string;
   var KeyCryptDisc1       : string;
   var KeyCryptDisc2       : string;
   var KeyCryptDisc3       : string;
   var KeyCryptResult : Boolean;

   begin
     //******************** �������� ������ **********************************************


     KeyCryptDisc1 := GetHardDiskSerial('c');
     KeyCryptDisc2 := EnCrypt (KeyCryptDisc1);
     KeyCryptDisc3 := IniFile.ReadString(IniSectionValue [1,1], IniSectionValue [2,2], IniSectionValue [2,3]);
     //showMessage(KeyCryptDisc3);
     if KeyCryptDisc2 = KeyCryptDisc3
      then

       Exit;

    //********* ���� ������ �� ������� �� �������� ������ ******************************


     Password := InputBox('AUTHORIZATION','���������� ������� ������','���������') ;
    // showMessage(Password);
    // ShowMessage(EnCrypt(Password));

    KeyCrypt:='lbvegaardnlbvegaardn';

    if EnCrypt(Password) = KeyCrypt then    KeyCryptResult := True
                                    else    KeyCryptResult := False;



    if KeyCryptResult then  IniFile.WriteString(IniSectionValue [1,1], IniSectionValue [2,2], KeyCryptDisc2)
                      else
                      begin
                        ShowMessage('�������� ������ ���������');
                        Application.Terminate;
                      end;





   end;


   //********************** ��������� ������ � INI  ������� **************************
 procedure IniFileCreate;
    begin
    IniSectionValue [1,1]:= 'FileLocation';
    IniSectionValue [1,2]:= 'LocationOBZOR';
    IniSectionValue [1,3]:= 'Z:\GRAND NEVA\2014\������ ������\'; //��� ������ �������� � ������ INI �����
    IniFile:=TIniFile.Create(ExtractFilePath(Application.ExeName)+ProgramName+'.ini');
    //******************* For Crypt Information **********************************
    IniSectionValue [2,2]:= 'Capasity Data';
    IniSectionValue [2,3]:= 'shhhdbbdb56';
    IniSectionValue [4,2]:= 'Object OBZOR Location';
    IniSectionValue [4,3]:= 'Z:\GRAND NEVA\2014\������ ������\';
    IniSectionValue [5,2]:= 'RF OBZOR Location';
    IniSectionValue [5,3]:= 'Z:\GRAND NEVA\2014\����� ��\';
    IniSectionValue [6,2]:= 'REGION OBZOR Location';
    IniSectionValue [6,3]:= 'Z:\GRAND NEVA\2014\����� ������\';
    IniSectionValue [7,2]:= 'DOCI Ocenschik Location';
    IniSectionValue [7,3]:= 'Z:\GRAND NEVA\2014\���� ��������\';
    IniSectionValue [8,2]:= 'Recvizity Ocenschik Location';
    IniSectionValue [8,3]:= 'Z:\GRAND NEVA\2014\���� ��������\';
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

