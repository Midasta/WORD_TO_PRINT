object Form4: TForm4
  Left = 0
  Top = 0
  Caption = 'EXCEL TO WORD'
  ClientHeight = 779
  ClientWidth = 916
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 916
    Height = 249
    Align = alTop
    TabOrder = 0
    object Label1: TLabel
      Left = 16
      Top = 59
      Width = 19
      Height = 13
      Caption = #1055#1086#1083
    end
    object Label2: TLabel
      Left = 16
      Top = 105
      Width = 80
      Height = 13
      Caption = #1044#1072#1090#1072' '#1088#1086#1078#1076#1077#1085#1080#1103
    end
    object ComboBox1: TComboBox
      Left = 16
      Top = 78
      Width = 56
      Height = 21
      TabOrder = 1
      Items.Strings = (
        ''
        #1052
        #1046)
    end
    object LabeledEdit1: TLabeledEdit
      Left = 16
      Top = 32
      Width = 191
      Height = 21
      EditLabel.Width = 23
      EditLabel.Height = 13
      EditLabel.Caption = #1060#1048#1054
      TabOrder = 0
    end
    object MaskEdit1: TMaskEdit
      Left = 16
      Top = 124
      Width = 120
      Height = 21
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      TabOrder = 2
      Text = '  .  .    '
    end
    object LabeledEdit2: TLabeledEdit
      Left = 16
      Top = 168
      Width = 193
      Height = 21
      EditLabel.Width = 31
      EditLabel.Height = 13
      EditLabel.Caption = #1040#1076#1088#1077#1089
      TabOrder = 3
    end
    object LabeledEdit3: TLabeledEdit
      Left = 16
      Top = 208
      Width = 193
      Height = 21
      EditLabel.Width = 78
      EditLabel.Height = 13
      EditLabel.Caption = #1057#1087#1077#1094#1080#1072#1083#1100#1085#1086#1089#1090#1100
      TabOrder = 4
    end
    object Button1: TButton
      Left = 344
      Top = 208
      Width = 75
      Height = 25
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100
      TabOrder = 5
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 425
      Top = 208
      Width = 75
      Height = 25
      Caption = #1054#1095#1080#1089#1090#1080#1090#1100
      TabOrder = 6
      OnClick = Button2Click
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 249
    Width = 916
    Height = 479
    Align = alClient
    TabOrder = 1
    object DBGrid1: TDBGrid
      Left = 1
      Top = 1
      Width = 914
      Height = 477
      Align = alClient
      DataSource = DataSource1
      Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgConfirmDelete, dgCancelOnExit, dgTitleClick, dgTitleHotTrack]
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'Tahoma'
      TitleFont.Style = []
      Columns = <
        item
          Expanded = False
          FieldName = 'ID'
          Visible = False
        end
        item
          Expanded = False
          FieldName = 'NAME_RU'
          Width = 150
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'POL'
          Width = 50
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'DOB'
          Title.Caption = #1044#1072#1090#1072' '#1088#1086#1078#1076'.'
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'ADRESS'
          Width = 250
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'SPEC'
          Width = 250
          Visible = True
        end>
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 728
    Width = 916
    Height = 51
    Align = alBottom
    TabOrder = 2
    object Button3: TButton
      Left = 632
      Top = 16
      Width = 99
      Height = 25
      Caption = #1069#1082#1089#1087#1086#1088#1090' '#1074' '#1042#1054#1056#1044
      TabOrder = 0
      OnClick = Button3Click
    end
    object Button4: TButton
      Left = 737
      Top = 16
      Width = 104
      Height = 25
      Caption = #1069#1082#1089#1087#1086#1088#1090' '#1074' '#1069#1050#1057#1045#1051#1068
      TabOrder = 1
      OnClick = Button4Click
    end
  end
  object ADOConnection1: TADOConnection
    Connected = True
    ConnectionString = 
      'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Rustam\De' +
      'sktop\TEST\Win32\Debug\Database.accdb;Persist Security Info=Fals' +
      'e'
    LoginPrompt = False
    Mode = cmShareDenyNone
    Provider = 'Microsoft.ACE.OLEDB.12.0'
    Left = 584
    Top = 32
  end
  object ADOQuery1: TADOQuery
    Active = True
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'Select * from Users')
    Left = 648
    Top = 40
    object ADOQuery1NAME_RU: TWideStringField
      DisplayLabel = #1060#1048#1054
      FieldName = 'NAME_RU'
      Size = 255
    end
    object ADOQuery1POL: TWideStringField
      DisplayLabel = #1055#1086#1083
      FieldName = 'POL'
      Size = 255
    end
    object ADOQuery1DOB: TDateTimeField
      DisplayLabel = #1044#1072#1090#1072' '#1088#1086#1078#1076#1077#1085#1080#1103
      FieldName = 'DOB'
    end
    object ADOQuery1ADRESS: TWideStringField
      DisplayLabel = #1040#1076#1088#1077#1089
      FieldName = 'ADRESS'
      Size = 255
    end
    object ADOQuery1SPEC: TWideStringField
      DisplayLabel = #1057#1087#1077#1094#1080#1072#1083#1100#1085#1086#1089#1090#1100
      FieldName = 'SPEC'
      Size = 255
    end
  end
  object DataSource1: TDataSource
    DataSet = ADOQuery1
    Left = 712
    Top = 48
  end
end
