object FrmAddObjToExpOtech: TFrmAddObjToExpOtech
  Left = 244
  Top = 247
  BorderStyle = bsDialog
  Caption = #1053#1086#1074#1099#1081' '#1086#1073#1098#1077#1082#1090
  ClientHeight = 341
  ClientWidth = 589
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object ButtonOK: TAdvGlowButton
    Left = 352
    Top = 292
    Width = 100
    Height = 33
    Caption = #1054#1050
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    TabOrder = 0
    OnClick = ButtonOKClick
    Appearance.ColorChecked = 16111818
    Appearance.ColorCheckedTo = 16367008
    Appearance.ColorDisabled = 15921906
    Appearance.ColorDisabledTo = 15921906
    Appearance.ColorDown = 16111818
    Appearance.ColorDownTo = 16367008
    Appearance.ColorHot = 16117985
    Appearance.ColorHotTo = 16372402
    Appearance.ColorMirror = clBtnFace
    Appearance.ColorMirrorHot = 16107693
    Appearance.ColorMirrorHotTo = 16775412
    Appearance.ColorMirrorDown = 16102556
    Appearance.ColorMirrorDownTo = 16768988
    Appearance.ColorMirrorChecked = 16102556
    Appearance.ColorMirrorCheckedTo = 16768988
    Appearance.ColorMirrorDisabled = 16102556
    Appearance.ColorMirrorDisabledTo = 15921906
  end
  object AdvGlowButton1: TAdvGlowButton
    Left = 468
    Top = 292
    Width = 100
    Height = 33
    Caption = #1054#1090#1084#1077#1085#1072
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    TabOrder = 1
    OnClick = AdvGlowButton1Click
    Appearance.ColorChecked = 16111818
    Appearance.ColorCheckedTo = 16367008
    Appearance.ColorDisabled = 15921906
    Appearance.ColorDisabledTo = 15921906
    Appearance.ColorDown = 16111818
    Appearance.ColorDownTo = 16367008
    Appearance.ColorHot = 16117985
    Appearance.ColorHotTo = 16372402
    Appearance.ColorMirror = clBtnFace
    Appearance.ColorMirrorHot = 16107693
    Appearance.ColorMirrorHotTo = 16775412
    Appearance.ColorMirrorDown = 16102556
    Appearance.ColorMirrorDownTo = 16768988
    Appearance.ColorMirrorChecked = 16102556
    Appearance.ColorMirrorCheckedTo = 16768988
    Appearance.ColorMirrorDisabled = 16102556
    Appearance.ColorMirrorDisabledTo = 15921906
  end
  object cxLabel2: TcxLabel
    Left = 16
    Top = 24
    Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077
  end
  object ComboBoxObjName: TcxLookupComboBox
    Left = 144
    Top = 24
    Properties.DropDownRows = 20
    Properties.KeyFieldNames = 'IDOBJ'
    Properties.ListColumns = <
      item
        Caption = #1050#1086#1076' '#1086#1073#1098#1077#1082#1090#1072
        MinWidth = 0
        Width = 0
        FieldName = 'IDOBJ'
      end
      item
        Caption = #1053#1072#1079#1074#1072#1085#1080#1077', '#1083#1072#1090#1080#1085#1089#1082#1086#1077
        FieldName = 'NAMEOBJ'
      end
      item
        Caption = #1053#1072#1079#1074#1072#1085#1080#1077', '#1088#1091#1089#1089#1082#1086#1077
        FieldName = 'NAMEOBJRUS'
      end>
    Properties.ListFieldIndex = 1
    Properties.ListOptions.ShowHeader = False
    Properties.ListSource = DataSource2
    EditValue = 0
    Style.BorderStyle = ebsOffice11
    Style.ButtonStyle = btsOffice11
    TabOrder = 3
    Width = 425
  end
  object cxLabel3: TcxLabel
    Left = 16
    Top = 64
    Caption = #1057#1086#1089#1090#1086#1103#1085#1080#1077'  '
    Properties.WordWrap = True
  end
  object cxLabel1: TcxLabel
    Left = 16
    Top = 226
    Caption = #1055#1088#1080#1084#1077#1095#1072#1085#1080#1077
    Properties.WordWrap = True
  end
  object SostMemo: TcxMemo
    Left = 144
    Top = 61
    TabOrder = 6
    Height = 47
    Width = 424
  end
  object MemoPrim: TcxMemo
    Left = 144
    Top = 227
    TabOrder = 7
    Height = 47
    Width = 424
  end
  object ButtonClearAll: TAdvGlowButton
    Left = 16
    Top = 292
    Width = 100
    Height = 33
    Caption = #1054#1095#1080#1089#1090#1080#1090#1100' '#1074#1089#1105
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    TabOrder = 8
    OnClick = ButtonClearAllClick
    Appearance.ColorChecked = 16111818
    Appearance.ColorCheckedTo = 16367008
    Appearance.ColorDisabled = 15921906
    Appearance.ColorDisabledTo = 15921906
    Appearance.ColorDown = 16111818
    Appearance.ColorDownTo = 16367008
    Appearance.ColorHot = 16117985
    Appearance.ColorHotTo = 16372402
    Appearance.ColorMirror = clBtnFace
    Appearance.ColorMirrorHot = 16107693
    Appearance.ColorMirrorHotTo = 16775412
    Appearance.ColorMirrorDown = 16102556
    Appearance.ColorMirrorDownTo = 16768988
    Appearance.ColorMirrorChecked = 16102556
    Appearance.ColorMirrorCheckedTo = 16768988
    Appearance.ColorMirrorDisabled = 16102556
    Appearance.ColorMirrorDisabledTo = 15921906
  end
  object LabelOchag: TcxLabel
    Left = 16
    Top = 124
    AutoSize = False
    Caption = #1057#1074#1077#1076#1077#1085#1080#1103' '#1086#1073' '#1086#1095#1072#1075#1077' ('#1084#1077#1089#1090#1086#1085#1072#1093#1086#1078#1076#1077#1085#1080#1077' '#1091#1095#1072#1089#1090#1082#1072', '#1082#1091#1083#1100#1090#1091#1088#1072', '#1089#1086#1088#1090')'
    Properties.WordWrap = True
    Height = 49
    Width = 129
  end
  object cxLabel13: TcxLabel
    Left = 16
    Top = 185
    AutoSize = False
    Caption = #1050#1086#1083#1080#1095#1077#1089#1090#1074#1086' '#1079#1072#1088#1072#1078#1105#1085#1085#1086#1081' '#1087#1088#1086#1076#1091#1082#1094#1080#1080' ('#1079#1072#1088#1072#1078#1105#1085#1085#1072#1103' '#1087#1083#1086#1097#1072#1076#1100')'
    Properties.WordWrap = True
    Height = 33
    Width = 201
  end
  object CalcEditEdIzm: TcxCalcEdit
    Left = 229
    Top = 189
    EditValue = 0
    Properties.ImmediatePost = True
    TabOrder = 12
    Width = 97
  end
  object cxLabel18: TcxLabel
    Left = 375
    Top = 190
    Caption = #1074' '#1077#1076#1080#1085#1080#1094#1072#1093' '#1080#1079#1084'.'
  end
  object ComboBoxEdIzm: TcxLookupComboBox
    Left = 470
    Top = 189
    Properties.DropDownRows = 20
    Properties.ImmediatePost = True
    Properties.KeyFieldNames = 'IDEDIZM'
    Properties.ListColumns = <
      item
        MinWidth = 0
        Width = 0
        FieldName = 'IDEDIZM'
      end
      item
        FieldName = 'NAMEEDIZM'
      end>
    Properties.ListFieldIndex = 1
    Properties.ListOptions.ShowHeader = False
    Properties.ListSource = DataSource10
    EditValue = 0
    Style.BorderStyle = ebsOffice11
    Style.ButtonStyle = btsOffice11
    TabOrder = 14
    Width = 96
  end
  object ButtonMemoSize: TcxButton
    Left = 548
    Top = 126
    Width = 19
    Height = 19
    Hint = #1048#1079#1084#1077#1085#1080#1090#1100' '#1088#1072#1079#1084#1077#1088' '#1086#1073#1083#1072#1089#1090#1080' '#1088#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1085#1080#1103
    Caption = '<>'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 15
    OnClick = ButtonMemoSizeClick
  end
  object MemoOchag: TcxMemo
    Left = 144
    Top = 125
    TabOrder = 10
    Height = 47
    Width = 406
  end
  object ADODataSet2: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from KAROBJECTS ORDER BY TYPEOBJ, NAMEOBJ;'
    CommandTimeout = 600
    Parameters = <>
    Left = 172
    Top = 305
  end
  object DataSource2: TDataSource
    DataSet = ADODataSet2
    Left = 200
    Top = 305
  end
  object ADODataSet10: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from EDIZM;'
    CommandTimeout = 600
    Parameters = <>
    Left = 240
    Top = 305
  end
  object DataSource10: TDataSource
    DataSet = ADODataSet10
    Left = 268
    Top = 305
  end
end
