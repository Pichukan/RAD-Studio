object Form2: TForm2
  Left = 0
  Top = 0
  Caption = 'INPUT PASSWORD'
  ClientHeight = 75
  ClientWidth = 273
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 16
    Width = 213
    Height = 19
    Caption = #1042#1074#1077#1076#1080#1090#1077' '#1087#1072#1088#1086#1083#1100' '#1082' '#1087#1088#1086#1075#1088#1072#1084#1084#1077
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Edit1: TEdit
    Left = 8
    Top = 46
    Width = 137
    Height = 21
    TabOrder = 0
    Text = 'Edit1'
  end
  object Button1: TButton
    Left = 160
    Top = 42
    Width = 65
    Height = 25
    Caption = 'GO'
    TabOrder = 1
    OnClick = Button1Click
  end
end
