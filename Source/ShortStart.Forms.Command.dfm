object CommandForm: TCommandForm
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = 'CommandForm'
  ClientHeight = 332
  ClientWidth = 560
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 13
    Width = 45
    Height = 13
    Caption = 'File name'
  end
  object Label2: TLabel
    Left = 16
    Top = 106
    Width = 48
    Height = 13
    Caption = 'Operation'
  end
  object Label3: TLabel
    Left = 16
    Top = 157
    Width = 55
    Height = 13
    Caption = 'Parameters'
  end
  object Label4: TLabel
    Left = 16
    Top = 213
    Width = 85
    Height = 13
    Caption = 'Working directory'
  end
  object Label5: TLabel
    Left = 16
    Top = 62
    Width = 41
    Height = 13
    Caption = 'Shortcut'
  end
  object FileNameEdit: TEdit
    Left = 16
    Top = 32
    Width = 489
    Height = 21
    TabOrder = 0
    Text = 'FileNameEdit'
  end
  object OperationEdit: TComboBox
    Left = 16
    Top = 125
    Width = 145
    Height = 21
    TabOrder = 3
    Text = 'OperationEdit'
    Items.Strings = (
      'open'
      'edit'
      'runas'
      'explore'
      'print')
  end
  object ParametersEdit: TEdit
    Left = 16
    Top = 176
    Width = 536
    Height = 21
    TabOrder = 4
    Text = 'ParametersEdit'
  end
  object WorkingDirectoryEdit: TEdit
    Left = 16
    Top = 232
    Width = 536
    Height = 21
    TabOrder = 5
    Text = 'WorkingDirectoryEdit'
  end
  object Button1: TButton
    Left = 511
    Top = 30
    Width = 42
    Height = 25
    Caption = '...'
    TabOrder = 2
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 384
    Top = 296
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 6
  end
  object Button3: TButton
    Left = 477
    Top = 296
    Width = 75
    Height = 25
    Cancel = True
    Caption = 'Cancel'
    ModalResult = 2
    TabOrder = 7
  end
  object ShortCutEdit: THotKey
    Left = 16
    Top = 81
    Width = 145
    Height = 19
    HotKey = 32833
    TabOrder = 1
  end
  object FileOpenDialog: TOpenDialog
    DefaultExt = 'exe'
    Filter = 'All excutables|*.exe;*.bat; *.cmd|Any file|*.*'
    Left = 392
    Top = 112
  end
end
