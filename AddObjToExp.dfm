object FrmAddObjToExp: TFrmAddObjToExp
  Left = 155
  Top = 390
  BorderStyle = bsDialog
  Caption = #1053#1086#1074#1099#1081' '#1086#1073#1098#1077#1082#1090
  ClientHeight = 236
  ClientWidth = 590
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object ButtonOK: TAdvGlowButton
    Left = 352
    Top = 186
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
    Top = 186
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
    Left = 117
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
    Width = 452
  end
  object cxLabel3: TcxLabel
    Left = 16
    Top = 63
    Caption = #1057#1086#1089#1090#1086#1103#1085#1080#1077'  '
    Properties.WordWrap = True
  end
  object cxLabel1: TcxLabel
    Left = 16
    Top = 122
    Caption = #1055#1088#1080#1084#1077#1095#1072#1085#1080#1077
    Properties.WordWrap = True
  end
  object SostMemo: TcxMemo
    Left = 118
    Top = 60
    TabOrder = 6
    Height = 47
    Width = 450
  end
  object MemoPrim: TcxMemo
    Left = 118
    Top = 123
    TabOrder = 7
    Height = 47
    Width = 450
  end
  object ButtonClearAll: TAdvGlowButton
    Left = 16
    Top = 186
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
  object ADODataSet2: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from KAROBJECTS ORDER BY TYPEOBJ, NAMEOBJ;'
    CommandTimeout = 600
    Parameters = <>
    Left = 17
    Top = 195
  end
  object DataSource2: TDataSource
    DataSet = ADODataSet2
    Left = 45
    Top = 195
  end
end
