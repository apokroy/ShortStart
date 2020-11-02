object AboutBox: TAboutBox
  Left = 200
  Top = 108
  BorderStyle = bsDialog
  Caption = 'About'
  ClientHeight = 223
  ClientWidth = 339
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object ScrollBox1: TScrollBox
    Left = 0
    Top = 0
    Width = 339
    Height = 169
    Align = alTop
    BorderStyle = bsNone
    Color = clWindow
    ParentColor = False
    TabOrder = 0
    object Comments: TLabel
      Left = 8
      Top = 112
      Width = 282
      Height = 41
      AutoSize = False
      Caption = 
        'Some icons by Yusuke Kamiyamane. Licensed under a Creative Commo' +
        'ns Attribution 3.0 License.'
      WordWrap = True
      IsControl = True
    end
    object Copyright: TLabel
      Left = 8
      Top = 80
      Width = 153
      Height = 13
      Caption = 'Copyright '#169' 2020 Alexey Pokroy'
      IsControl = True
    end
    object ProgramIcon: TImage
      Left = 8
      Top = 8
      Width = 48
      Height = 48
      Center = True
      Stretch = True
      IsControl = True
    end
    object Version: TLabel
      Left = 62
      Top = 43
      Width = 53
      Height = 13
      Caption = 'Version 1.0'
      IsControl = True
    end
    object ProductName: TLabel
      Left = 62
      Top = 8
      Width = 90
      Height = 19
      Caption = 'Short Start'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
      IsControl = True
    end
  end
  object OKButton: TButton
    Left = 256
    Top = 188
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 1
  end
end
