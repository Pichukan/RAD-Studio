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
    Label1: TLabel;


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
     ReplaceText [1]:='#1';        //���������� �� ������� ����� � WORD
     ReplaceText [2]:='#2';        //���������� �� ������� ����� � WORD
     ReplaceText [3]:='#3';        //���������� �� ������� ����� � WORD
     ReplaceText [4]:='#4';        //���������� �� ������� ����� � WORD
     ReplaceText [5]:='#5';        //���������� �� ������� ����� � WORD

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
procedure InsertPictureWord;
 var
 ReplaceText   : array [1..5] of string;
 DirPictureDoc : string;
 i             : Integer;

begin
  DirPictureDoc:= DIRFile+'����.docx';
  BookPicture:=WRD.Documents.Open(DirPictureDoc);
 // ShowMessage('���� ���� ������');

   ReplaceText [1]:='##1';
   ReplaceText [2]:='##2';
   ReplaceText [3]:='##3';
   ReplaceText [4]:='';

   BookPicture.Range.InlineShapes.Item(1).Range.CopyAsPicture;
   MyRange2 := FindInDoc(Book, ReplaceText [1]);

    if  VarIsClear(MyRange2) then begin
        ShowMessage('����� ##1 �� ������.');
        Exit;
    end;

     MyRange2.Paste;

     for i:=2 to 4  do
       begin
         BookPicture.Range.InlineShapes.Item(i).Range.CopyAsPicture;
         MyRange2 := FindInDoc(Book, ReplaceText [2]);

         if VarIsClear(MyRange2) then begin
            ShowMessage('����� ##2 �� ������.');
            Exit;
         end;

         MyRange2.Paste;
         if i<4 then  MyRange2.InsertAfter(ReplaceText [2])  ;

       end;

    for i:=5 to 100  do
       begin
         MyRange2 := FindInDoc(Book, ReplaceText [3]);

         if VarIsClear(MyRange2) then begin
            ShowMessage('����� ##3 �� ������.');
            Exit;
         end;

         try
           BookPicture.Range.InlineShapes.Item(i).Range.CopyAsPicture;
         except
           break;
         end;
         MyRange2.Paste;
         MyRange2.InsertAfter(ReplaceText [3])  ;
       end;

    //  ShowMessage(IntToStr(i));
      MyRange2 := FindInDoc(Book, ReplaceText [3]);
      MyRange2.text:=ReplaceText [4];
  // MyRange2.Paste;
  //    ShowMessage('��� ');


end;

//**** ��������� ���� �� ������� EXCEL �������� ������               *****
//**** ��������� �������� WORD � ������� �� �������� ���������       *****
//**** � ��������� � �������� WORD ������                            *****
 procedure InsertObzor;
 var
    DirObzor,ReplaceTextObzor : string;

 begin
      MyWorkSheet2:=MyBook.Sheets['����'];
      RangeObzor:=MyWorkSheet2.Range['b22'];
      ObzorValue:= RangeObzor.Value;
    //ShowMessage(vartostr(ObzorValue));
      DirObzor:= 'Z:\GRAND NEVA\2014\������ ������\'+ObzorValue+' �����.docx';
      BookObzor:=WRD.Documents.Open(DirObzor);
      BookObzor.Range.Copy;
      ReplaceTextObzor:='###1';
      MyRange2 := FindInDoc(Book, ReplaceTextObzor);

         if VarIsClear(MyRange2) then begin
            ShowMessage('����� ###1 ReplaceTextObzor �� ������.');
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
    BookAnalog:=WRD.Documents.Open(DirAnalogDoc);
    BookAnalog.Range.Copy;
    ReplaceTextAnalog:='####1';
    MyRange2 := FindInDoc(Book, ReplaceTextAnalog);

         if VarIsClear(MyRange2) then begin
            ShowMessage('����� ####1 ReplaceTextAnalog �� ������.');
            Exit;
         end;

    MyRange2.Paste;

 end;

//**** ��������� ��������� ������� �������������� �� ��������� WORD      ****
//**** � �������� ������                                                 ****
  procedure InsertLocation;
  var
       i: Integer;
       DirLocationDoc  : string;
       ReplaceTextLoc : array [1..3] of string;

  begin
       ReplaceTextLoc [1]:='#####1';
       ReplaceTextLoc [2]:='#####2';
       ReplaceTextLoc [3]:='#####3';


     DirLocationDoc:= DIRFile+'�����.docx';
     BookLocation:=WRD.Documents.Open(DirLocationDoc);

       for i:=1 to 3  do
       begin
         BookLocation.Range.InlineShapes.Item(i).Range.CopyAsPicture;
         MyRange2 := FindInDoc(Book, ReplaceTextLoc [i]);
         if VarIsClear(MyRange2) then begin
            ShowMessage('����� ####1 Location �� ������.');
            Exit;
         end;
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
  //  BookObzor.Close;
    BookPicture.Close;
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

procedure TForm1.Button2Click(Sender: TObject);
begin

    StartWORD;

   if not OpenDialog1.Execute then Exit;  //��� ����������� ������ ������ �����, � ���� ������������ ����� "Cancel", �� �������

    WRD.visible:=true;     //������ ���� ������� or not
    Book:= WRD.Documents.Open(OpenDialog1.FileName);
  // WRD.visible:=false;     //������ ���� ������� or not

    DIRFName := OpenDialog1.FileName;    //���������� ���������� ��������� �����
  //  ShowMessage(DIRFName);             // � ������ � �����������


    DIRName:= DIRDetect;    // ������������ ���������� � ��������� �����
    ShowMessage(DIRName);   // ��� ����������  �� � ������


    ShowMessage(DIRFileDetect);    // ������������ ���������� ����� ��� �������� �����
    DIRFile:= DIRFileDetect;       //� ������������� ����������

    StartExcel;
    ActExcelOpenDoc;             //��������� � ��������� EXCEL �� �������� ����������
    ExeMacros;
    EXC.visible:=true;       //������ Excel �������
 // EXC.visible:=false;      //������ Excel ���������

    TableAsPicturePaste(Book);
    InsertPictureWord;
  //  InsertObzor ;
    InsertAnalogi;
    InsertLocation;
    MacroSli;
    CloseWordDocs;
    CloseExcel;

  WRD.visible:=true;     //������ ���� �������
  ShowMessage('program '+ ProgramName + ' end');
  //Book.Range(1,10).Paste;



end;


begin
   ProgramName:='Blevantone 1.0';
  // TForm1.Caption:= ProgramName;





 end.

