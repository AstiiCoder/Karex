object FrmPeriodExp: TFrmPeriodExp
  Left = 422
  Top = 468
  BorderStyle = bsDialog
  Caption = #1047#1072#1087#1088#1086#1089' '#1087#1072#1088#1072#1084#1077#1090#1088#1086#1074
  ClientHeight = 285
  ClientWidth = 555
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object ButtonOK: TAdvGlowButton
    Left = 37
    Top = 238
    Width = 100
    Height = 33
    Caption = #1054#1050
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
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
  object ButtonCancel: TAdvGlowButton
    Left = 166
    Top = 237
    Width = 100
    Height = 33
    Caption = #1054#1090#1084#1077#1085#1072
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    OnClick = ButtonCancelClick
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
  object cxGroupBox1: TcxGroupBox
    Left = 17
    Top = 14
    Caption = #1055#1077#1088#1080#1086#1076' '#1087#1088#1086#1074#1077#1076#1077#1085#1080#1103' '#1101#1082#1089#1087#1077#1088#1090#1080#1079
    Style.BorderStyle = ebsFlat
    TabOrder = 2
    Height = 65
    Width = 271
    object cxLabel1: TcxLabel
      Left = 10
      Top = 25
      Caption = 'c:'
      Transparent = True
    end
    object Date1: TcxDateEdit
      Left = 31
      Top = 25
      EditValue = 0d
      Properties.DateButtons = [btnToday]
      Properties.ImmediatePost = True
      Style.BorderColor = clBtnFace
      Style.BorderStyle = ebsOffice11
      Style.ButtonStyle = btsOffice11
      TabOrder = 1
      Width = 80
    end
    object LabelDate2: TcxLabel
      Left = 141
      Top = 25
      Caption = #1087#1086':'
      Transparent = True
    end
    object Date2: TcxDateEdit
      Left = 173
      Top = 25
      EditValue = 0d
      Properties.DateButtons = [btnToday]
      Properties.ImmediatePost = True
      Style.BorderColor = clBtnFace
      Style.BorderStyle = ebsOffice11
      Style.ButtonStyle = btsOffice11
      TabOrder = 3
      Width = 80
    end
  end
  object RadioGroupParamVibor: TcxRadioGroup
    Left = 17
    Top = 92
    Caption = #1055#1072#1088#1072#1084#1077#1090#1088#1099' '#1086#1090#1073#1086#1088#1072
    Properties.ImmediatePost = True
    Properties.Items = <
      item
        Caption = #1074#1089#1077' '#1101#1082#1089#1087#1077#1088#1090#1080#1079#1099
      end
      item
        Caption = #1090#1086#1083#1100#1082#1086' '#1089' '#1085#1072#1081#1076#1077#1085#1085#1099#1084#1080' '#1086#1073#1098#1077#1082#1090#1072#1084#1080
      end
      item
        Caption = #1087#1086' '#1085#1086#1084#1077#1088#1091' '#1087#1088#1086#1090#1086#1082#1086#1083#1072
      end>
    Properties.OnChange = RadioGroupParamViborPropertiesChange
    Properties.OnEditValueChanged = RadioGroupParamViborPropertiesEditValueChanged
    Style.BorderStyle = ebsFlat
    TabOrder = 3
    Height = 128
    Width = 271
  end
  object RadioGroupObj: TcxRadioGroup
    Left = 305
    Top = 14
    Caption = #1053#1072#1081#1076#1077#1085#1085#1099#1077' '#1086#1073#1098#1077#1082#1090#1099
    Properties.Items = <
      item
        Caption = #1082#1072#1088#1072#1085#1090#1080#1085#1085#1099#1077
      end
      item
        Caption = #1085#1077#1082#1072#1088#1072#1085#1090#1080#1085#1085#1099#1077', '#1086#1090#1089#1091#1090#1089#1090#1074#1091#1102#1097#1080#1077' '#1074' '#1056#1060
      end
      item
        Caption = #1085#1077#1082#1072#1088#1072#1085#1090#1080#1085#1085#1099#1077', '#1087#1088#1086#1095#1080#1077
      end>
    ItemIndex = 0
    Style.BorderStyle = ebsFlat
    TabOrder = 4
    Height = 112
    Width = 232
  end
  object RadioGroupImport: TcxRadioGroup
    Left = 305
    Top = 139
    Caption = #1055#1088#1086#1080#1089#1093#1086#1078#1076#1077#1085#1080#1077' '#1086#1073#1088#1072#1079#1094#1072
    Properties.Items = <
      item
        Caption = #1080#1084#1087#1086#1088#1090#1085#1099#1077
      end
      item
        Caption = #1086#1090#1077#1095#1077#1089#1090#1074#1077#1085#1085#1099#1077
      end>
    ItemIndex = 0
    Style.BorderStyle = ebsFlat
    TabOrder = 5
    Height = 82
    Width = 232
  end
  object CalcEditCode: TcxCalcEdit
    Left = 173
    Top = 189
    EditValue = 0.000000000000000000
    Properties.ImmediatePost = True
    TabOrder = 6
    Visible = False
    Width = 97
  end
end
