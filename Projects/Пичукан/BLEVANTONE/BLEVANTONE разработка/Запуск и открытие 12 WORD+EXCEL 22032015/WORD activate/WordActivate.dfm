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
    Left = 96
    Top = 143
    Width = 277
    Height = 18
    Caption = #1048#1085#1076#1080#1082#1072#1094#1080#1103' '#1074#1099#1087#1086#1083#1085#1077#1085#1080#1103' '#1087#1088#1086#1075#1088#1072#1084#1084#1099
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label2: TLabel
    Left = 56
    Top = 188
    Width = 47
    Height = 16
    Caption = 'Label2'
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Bookman Old Style'
    Font.Style = [fsBold, fsItalic]
    ParentFont = False
  end
  object Button2: TButton
    Left = 56
    Top = 8
    Width = 393
    Height = 41
    Caption = 'Start Program'
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'Bookman Old Style'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 0
    OnClick = Button2Click
  end
  object ProgressBar1: TProgressBar
    Left = 56
    Top = 165
    Width = 393
    Height = 17
    MarqueeInterval = 1
    TabOrder = 1
  end
  object OpenDialog1: TOpenDialog
    Left = 8
    Top = 88
  end
end
