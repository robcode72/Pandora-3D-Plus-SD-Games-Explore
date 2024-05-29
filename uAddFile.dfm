object frmAddFile: TfrmAddFile
  Left = 0
  Top = 0
  ActiveControl = edtDEFAULT_DISPLAY_NAME
  BorderStyle = bsDialog
  Caption = 'Add new file'
  ClientHeight = 530
  ClientWidth = 769
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  Position = poMainFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 769
    Height = 489
    Align = alClient
    BevelKind = bkSoft
    TabOrder = 0
    object Label1: TLabel
      Left = 48
      Top = 222
      Width = 54
      Height = 13
      Caption = 'EMULATOR'
    end
    object lblGenre: TLabel
      Left = 199
      Top = 222
      Width = 33
      Height = 13
      Caption = 'GENRE'
    end
    object Label2: TLabel
      Left = 48
      Top = 364
      Width = 89
      Height = 13
      Caption = 'DEFAULT_CONFIG'
    end
    object Label3: TLabel
      Left = 208
      Top = 364
      Width = 116
      Height = 13
      Caption = 'SUPPORTED_FEATURES'
    end
    object Label4: TLabel
      Left = 48
      Top = 416
      Width = 31
      Height = 13
      Caption = 'FLAG1'
    end
    object Label5: TLabel
      Left = 240
      Top = 416
      Width = 31
      Height = 13
      Caption = 'FLAG3'
    end
    object Label6: TLabel
      Left = 145
      Top = 416
      Width = 31
      Height = 13
      Caption = 'FLAG2'
    end
    object edtDEFAULT_DISPLAY_NAME: TLabeledEdit
      Left = 248
      Top = 187
      Width = 433
      Height = 21
      EditLabel.Width = 128
      EditLabel.Height = 13
      EditLabel.Caption = 'DEFAULT_DISPLAY_NAME '
      TabOrder = 0
    end
    object cbEMULATOR: TComboBox
      Left = 48
      Top = 241
      Width = 145
      Height = 21
      Style = csDropDownList
      TabOrder = 1
    end
    object edtCOIN_DURATION: TLabeledEdit
      Left = 48
      Top = 289
      Width = 73
      Height = 21
      EditLabel.Width = 133
      EditLabel.Height = 13
      EditLabel.Caption = 'COIN_DURATION '#8211' (0-250)'
      MaxLength = 3
      NumbersOnly = True
      TabOrder = 2
      Text = '0'
    end
    object cbGENRE: TComboBox
      Left = 199
      Top = 241
      Width = 145
      Height = 21
      Style = csDropDownList
      TabOrder = 3
    end
    object edtLoadTime: TLabeledEdit
      Left = 48
      Top = 337
      Width = 73
      Height = 21
      EditLabel.Width = 117
      EditLabel.Height = 13
      EditLabel.Caption = 'LOAD_TIME '#8211' (0-50000)'
      MaxLength = 5
      NumbersOnly = True
      TabOrder = 4
      Text = '0'
    end
    object edtROM_NAME: TLabeledEdit
      Left = 48
      Top = 187
      Width = 184
      Height = 21
      EditLabel.Width = 57
      EditLabel.Height = 13
      EditLabel.Caption = 'ROM_NAME'
      Enabled = False
      ReadOnly = True
      TabOrder = 5
    end
    object cbDEFAULT_CONFIG: TComboBox
      Left = 48
      Top = 383
      Width = 145
      Height = 21
      ItemIndex = 0
      TabOrder = 6
      Text = 'NULL'
      Items.Strings = (
        'NULL'
        '4_.zip'
        '4_2.zip'
        '4_3.zip'
        '_.zip'
        '_2.zip'
        '_3.zip')
    end
    object cbSUPPORTED_FEATURES: TComboBox
      Left = 208
      Top = 383
      Width = 145
      Height = 21
      ItemIndex = 5
      TabOrder = 7
      Text = '28[savestate]'
      Items.Strings = (
        '4[exit only]'
        '5'
        '20[exit only]'
        '21'
        '23[exit only]'
        '28[savestate]'
        '29'
        '30'
        '31[savestate]')
    end
    object edtCOIN_INSERT_MODE: TLabeledEdit
      Left = 359
      Top = 383
      Width = 106
      Height = 21
      EditLabel.Width = 103
      EditLabel.Height = 13
      EditLabel.Caption = 'COIN_INSERT_MODE'
      MaxLength = 3
      NumbersOnly = True
      TabOrder = 8
      Text = '0'
    end
    object cbFLAG1: TComboBox
      Left = 48
      Top = 433
      Width = 89
      Height = 21
      ItemIndex = 0
      TabOrder = 9
      Text = 'NULL'
      Items.Strings = (
        'NULL'
        '0'
        '1[n64]')
    end
    object cbFLAG2: TComboBox
      Left = 145
      Top = 433
      Width = 89
      Height = 21
      ItemIndex = 0
      TabOrder = 10
      Text = 'NULL'
      Items.Strings = (
        'NULL')
    end
    object cbFLAG3: TComboBox
      Left = 240
      Top = 433
      Width = 89
      Height = 21
      ItemIndex = 0
      TabOrder = 11
      Text = 'NULL'
      Items.Strings = (
        'NULL')
    end
    object chkENTERNO: TCheckBox
      Left = 338
      Top = 438
      Width = 65
      Height = 17
      Caption = 'EXTERNA'
      TabOrder = 12
    end
    object GroupBox1: TGroupBox
      Left = 8
      Top = 9
      Width = 745
      Height = 144
      Caption = 'Real Path to file'
      TabOrder = 13
      object sbOpenFile: TSpeedButton
        Left = 572
        Top = 63
        Width = 23
        Height = 22
        Caption = '...'
        OnClick = sbOpenFileClick
      end
      object SpeedButton3: TSpeedButton
        Left = 572
        Top = 23
        Width = 23
        Height = 22
        Caption = '...'
        OnClick = SpeedButton3Click
      end
      object SpeedButton4: TSpeedButton
        Left = 3
        Top = 63
        Width = 23
        Height = 22
        Hint = 'Past clipboard'
        ImageIndex = 2
        Images = frmMain.IlComunes
        Flat = True
        ParentShowHint = False
        ShowHint = True
        OnClick = SpeedButton4Click
      end
      object edtVideoFile: TLabeledEdit
        Left = 137
        Top = 63
        Width = 433
        Height = 21
        EditLabel.Width = 86
        EditLabel.Height = 13
        EditLabel.Caption = 'VIDEO_FILE(MP4)'
        TabOrder = 0
      end
      object edtFILE_NAME: TLabeledEdit
        Left = 137
        Top = 23
        Width = 433
        Height = 21
        EditLabel.Width = 50
        EditLabel.Height = 13
        EditLabel.Caption = 'ROM_FILE'
        TabOrder = 1
      end
      object rdFileType: TRadioGroup
        Left = 137
        Top = 90
        Width = 429
        Height = 36
        Caption = 'File type'
        Columns = 2
        ItemIndex = 1
        Items.Strings = (
          'URL from YouTube'
          'MP4 File')
        TabOrder = 2
      end
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 489
    Width = 769
    Height = 41
    Align = alBottom
    TabOrder = 1
    object SpeedButton1: TSpeedButton
      Left = 567
      Top = 3
      Width = 97
      Height = 33
      Caption = 'Save'
      OnClick = SpeedButton1Click
    end
    object SpeedButton2: TSpeedButton
      Left = 670
      Top = 3
      Width = 97
      Height = 33
      Caption = 'Cancel'
      OnClick = SpeedButton2Click
    end
  end
  object OpenDialog: TOpenDialog
    Filter = '*.zip |zip'
    Left = 537
    Top = 317
  end
  object qGames_Insert: TFDQuery
    Connection = frmMain.FDConnection1
    SQL.Strings = (
      
        'INSERT INTO info (CREATE_ID, ROM_NAME, DEFAULT_DISPLAY_NAME,EMUL' +
        'ATOR,GENRE, LOAD_TIME, COIN_DURATION, COIN_INSERT_MODE, DEFAULT_' +
        'CONFIG, SUPPORT_FEATURES, FLAG1, FLAG2, FLAG3, EXTERNAL )'
      
        'values(:createid, :romname, :def_display_name, :emulator, :genre' +
        ', :load_time, :coin_duration, :coin_insert_mode, :default_conf, ' +
        ':suport_f, :flag1, :flag2, :flag3, :externo);'
      ''
      
        'INSERT INTO option (ROM_NAME, HIDE, TOP_ID, CURRENT_CONFIG, TARG' +
        'ET_CONFIG, PLAY_ID, KEY_COMBINATION)'
      'values(:romname, 0, 0, NULL, NULL, 0, 0);'
      ''
      'INSERT INTO display_name (ROM_NAME, LANGUAGE)'
      'values(:romname, "EN");')
    Left = 608
    Top = 168
    ParamData = <
      item
        Name = 'CREATEID'
        DataType = ftInteger
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'ROMNAME'
        DataType = ftString
        ParamType = ptInput
        Size = 100
        Value = Null
      end
      item
        Name = 'DEF_DISPLAY_NAME'
        DataType = ftString
        ParamType = ptInput
        Size = 100
        Value = Null
      end
      item
        Name = 'EMULATOR'
        DataType = ftInteger
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'GENRE'
        DataType = ftInteger
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'LOAD_TIME'
        DataType = ftInteger
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'COIN_DURATION'
        DataType = ftInteger
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'COIN_INSERT_MODE'
        DataType = ftInteger
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'DEFAULT_CONF'
        DataType = ftString
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'SUPORT_F'
        DataType = ftInteger
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'FLAG1'
        DataType = ftString
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'FLAG2'
        DataType = ftString
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'FLAG3'
        DataType = ftString
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'EXTERNO'
        DataType = ftInteger
        ParamType = ptInput
        Value = Null
      end>
  end
  object qGames_Update: TFDQuery
    Connection = frmMain.FDConnection1
    SQL.Strings = (
      'BEGIN;'
      'UPDATE info'
      '   SET DEFAULT_DISPLAY_NAME =:displayname,'
      '       EMULATOR = :emu,'
      '       GENRE = :gen,'
      '       LOAD_TIME = :loadtime,'
      '       COIN_DURATION = :coinduration,'
      '       DEFAULT_CONFIG = :defaultconfig,'
      '       SUPPORT_FEATURES = :supportfeatures,'
      '       COIN_INSERT_MODE = :coininsertmode,'
      '       FLAG1 = :flag1,'
      '       FLAG2 = :flag2,'
      '       FLAG3 = :flag3'
      ''
      ' WHERE ROM_NAME = :romname;'
      ''
      'COMMIT;')
    Left = 696
    Top = 168
    ParamData = <
      item
        Name = 'DISPLAYNAME'
        DataType = ftString
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'EMU'
        DataType = ftInteger
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'GEN'
        DataType = ftInteger
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'LOADTIME'
        DataType = ftInteger
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'COINDURATION'
        DataType = ftInteger
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'DEFAULTCONFIG'
        DataType = ftString
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'SUPPORTFEATURES'
        DataType = ftInteger
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'COININSERTMODE'
        DataType = ftInteger
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'FLAG1'
        DataType = ftString
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'FLAG2'
        DataType = ftString
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'FLAG3'
        DataType = ftString
        ParamType = ptInput
        Value = Null
      end
      item
        Name = 'ROMNAME'
        DataType = ftString
        ParamType = ptInput
        Value = Null
      end>
  end
  object qGet_Rec_Num: TFDQuery
    Connection = frmMain.FDConnection1
    SQL.Strings = (
      'SELECT * FROM info'
      'ORDER BY CREATE_ID DESC LIMIT 1;')
    Left = 456
    Top = 232
    object qGet_Rec_NumCREATE_ID: TIntegerField
      FieldName = 'CREATE_ID'
      Origin = 'CREATE_ID'
    end
  end
end
