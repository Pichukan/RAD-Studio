object Form1: TForm1
  Left = 0
  Top = 0
  BorderStyle = bsSingle
  Caption = 'Program BLEVANTONE (ver. 1.0)'
  ClientHeight = 235
  ClientWidth = 492
  Color = clBtnFace
  Font.Charset = RUSSIAN_CHARSET
  Font.Color = clWindowText
  Font.Height = -16
  Font.Name = 'Bookman Old Style'
  Font.Style = [fsBold]
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 19
  object Label1: TLabel
    Left = 120
    Top = 143
    Width = 208
    Height = 16
    Caption = #1048#1085#1076#1080#1082#1072#1094#1080#1103' '#1074#1099#1087#1086#1083#1085#1077#1085#1080#1103' '#1087#1088#1086#1075#1088#1072#1084#1084#1099
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Button2: TButton
    Left = 56
    Top = 8
    Width = 393
    Height = 41
    Caption = 'Start Program'
    TabOrder = 0
    OnClick = Button2Click
  end
  object OpenDialog1: TOpenDialog
    Left = 8
    Top = 88
  end
end
