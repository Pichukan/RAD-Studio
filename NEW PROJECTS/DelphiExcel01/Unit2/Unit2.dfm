object Form2: TForm2
  Left = 0
  Top = 0
  Caption = 'Form2'
  ClientHeight = 252
  ClientWidth = 492
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object CreateButton: TButton
    Left = 360
    Top = 8
    Width = 113
    Height = 25
    Caption = #1057#1086#1079#1076#1072#1090#1100' '#1069#1076#1080#1090#1099
    TabOrder = 0
    Visible = False
    OnClick = CreateButtonClick
  end
  object DelButton: TButton
    Left = 360
    Top = 39
    Width = 113
    Height = 25
    Caption = #1059#1076#1072#1083#1080#1090#1100' '#1069#1076#1080#1090#1099
    TabOrder = 1
    Visible = False
    OnClick = DelButtonClick
  end
end
