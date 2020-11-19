object FrmDialog2: TFrmDialog2
  Left = 587
  Top = 320
  BorderStyle = bsDialog
  Caption = #1052#1072#1089#1090#1077#1088' '#1087#1088#1086#1074#1077#1076#1077#1085#1080#1103' '#1082#1072#1088#1072#1085#1090#1080#1085#1085#1086#1081' '#1101#1082#1089#1087#1077#1088#1090#1080#1079#1099
  ClientHeight = 414
  ClientWidth = 636
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object DialogWizard: TPageControl
    Left = 0
    Top = 33
    Width = 636
    Height = 328
    ActivePage = TabSheet1
    Align = alTop
    Style = tsButtons
    TabOrder = 0
    object TabSheet1: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'1'
      object LabelDlgName: TcxLabel
        Left = 20
        Top = 20
        Caption = #1056#1077#1079#1091#1083#1100#1090#1072#1090#1099' '#1082#1072#1088#1072#1085#1090#1080#1085#1085#1086#1081' '#1101#1082#1089#1087#1077#1088#1090#1080#1079#1099
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object cxLabel27: TcxLabel
        Left = 20
        Top = 61
        Caption = #1060'.'#1048'.'#1054'. '#1089#1087#1077#1094#1080#1072#1083#1080#1089#1090#1072' '#1051#1072#1073'., '#1087#1088#1086#1074#1086#1076#1103#1097#1077#1075#1086' '#1101#1082#1089#1087#1077#1088#1090#1080#1079#1091
      end
      object cxLabel28: TcxLabel
        Left = 20
        Top = 102
        Caption = #1044#1086#1083#1078#1085#1086#1089#1090#1100
      end
      object TextEditDolzhnSpets: TcxTextEdit
        Left = 304
        Top = 98
        Enabled = False
        TabOrder = 1
        Width = 299
      end
      object ComboBoxSpetsialist: TcxLookupComboBox
        Left = 304
        Top = 59
        Properties.DropDownRows = 20
        Properties.ImmediatePost = True
        Properties.KeyFieldNames = 'IDSPETSIALIST'
        Properties.ListColumns = <
          item
            Caption = #1050#1086#1076' '#1089#1087#1077#1094#1080#1072#1083#1080#1089#1090#1072
            MinWidth = 0
            Width = 0
            FieldName = 'IDSPETSIALIST'
          end
          item
            Caption = #1060'.'#1048'.'#1054'.'
            FieldName = 'NAMESPETSIALIST'
          end
          item
            Caption = #1044#1086#1083#1078#1085#1086#1089#1090#1100
            FieldName = 'DOLZHNSPETSIALIST'
          end>
        Properties.ListFieldIndex = 1
        Properties.ListOptions.ShowHeader = False
        Properties.ListSource = DataSource1
        Properties.OnEditValueChanged = ComboBoxSpetsialistPropertiesEditValueChanged
        EditValue = 0
        Style.BorderStyle = ebsOffice11
        Style.ButtonStyle = btsOffice11
        TabOrder = 0
        Width = 299
      end
    end
    object TabSheet2: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'2'
      ImageIndex = 1
      object cxLabel1: TcxLabel
        Left = 20
        Top = 20
        Caption = #1055#1086#1089#1090#1091#1087#1080#1074#1096#1080#1081' '#1086#1073#1088#1072#1079#1077#1094
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object RadioGroupImport: TcxRadioGroup
        Left = 20
        Top = 51
        Properties.Items = <
          item
            Caption = #1080#1084#1087#1086#1088#1090#1085#1072#1103' '#1087#1088#1086#1076#1091#1082#1094#1080#1103
          end
          item
            Caption = #1086#1090#1077#1095#1077#1089#1090#1074#1077#1085#1085#1072#1103' '#1087#1088#1086#1076#1091#1082#1094#1080#1103
          end>
        Properties.OnEditValueChanged = RadioGroupImportPropertiesEditValueChanged
        ItemIndex = 0
        Style.BorderStyle = ebsOffice11
        TabOrder = 1
        Height = 98
        Width = 189
      end
    end
    object TabSheet3: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'3'
      ImageIndex = 2
      object LabelNameClient: TcxLabel
        Left = 18
        Top = 61
        Caption = #1055#1088#1086#1090#1086#1082#1086#1083' '#8470
      end
      object cxLabel5: TcxLabel
        Left = 20
        Top = 20
        Caption = #1042#1074#1086#1076' '#1076#1072#1085#1085#1099#1093' '#1087#1086' '#1087#1088#1086#1090#1086#1082#1086#1083#1091
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object CalcEditProtokolNum: TcxCalcEdit
        Left = 102
        Top = 61
        EditValue = 0.000000000000000000
        Properties.ImmediatePost = True
        Properties.OnEditValueChanged = CalcEditProtokolNumPropertiesEditValueChanged
        TabOrder = 2
        Width = 97
      end
      object cxLabel8: TcxLabel
        Left = 210
        Top = 62
        Caption = #1086#1090
      end
      object TextEditDateProtokol: TcxTextEdit
        Left = 236
        Top = 61
        Enabled = False
        TabOrder = 4
        Width = 97
      end
      object TextEditIDExp: TcxTextEdit
        Left = 348
        Top = 61
        Enabled = False
        TabOrder = 5
        Visible = False
        Width = 37
      end
    end
    object TabSheet4: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'4'
      ImageIndex = 3
      object cxLabel6: TcxLabel
        Left = 20
        Top = 20
        Caption = #1050#1072#1088#1072#1085#1090#1080#1085#1085#1099#1077' '#1086#1073#1098#1077#1082#1090#1099
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object cxRadioGroup3: TcxRadioGroup
        Left = 11
        Top = 37
        Properties.Items = <
          item
            Caption = #1085#1077' '#1086#1073#1085#1072#1088#1091#1078#1077#1085#1099
          end
          item
            Caption = #1086#1073#1085#1072#1088#1091#1078#1077#1085#1099
          end>
        ItemIndex = 0
        Style.BorderColor = clNone
        Style.BorderStyle = ebsNone
        TabOrder = 1
        Height = 75
        Width = 145
      end
    end
    object TabSheet5: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'5'
      ImageIndex = 4
      object cxLabel14: TcxLabel
        Left = 20
        Top = 20
        Caption = #1054#1073#1085#1072#1088#1091#1078#1077#1085#1085#1099#1077' '#1082#1072#1088#1072#1085#1090#1080#1085#1085#1099#1077' '#1086#1073#1098#1077#1082#1090#1099
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object ListViewKarObj: TcxListView
        Left = 19
        Top = 54
        Width = 584
        Height = 67
        Columns = <
          item
            Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077
            Width = 200
          end
          item
            Caption = #1057#1086#1089#1090#1086#1103#1085#1080#1077
            Width = 250
          end
          item
            Caption = #1055#1088#1080#1084#1077#1095#1072#1085#1080#1077
          end
          item
            Caption = #1050#1086#1076' '#1086#1073#1098#1077#1082#1090#1072
            Width = 0
          end>
        ShowWorkAreas = True
        Style.BorderStyle = cbsOffice11
        Style.LookAndFeel.Kind = lfOffice11
        StyleDisabled.LookAndFeel.Kind = lfOffice11
        StyleFocused.LookAndFeel.Kind = lfOffice11
        StyleHot.LookAndFeel.Kind = lfOffice11
        TabOrder = 1
        ViewStyle = vsReport
      end
      object ButtonAddKarObj: TAdvGlowButton
        Left = 20
        Top = 255
        Width = 102
        Height = 33
        Caption = #1044#1086#1073#1072#1074#1080#1090#1100
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        NotesFont.Charset = DEFAULT_CHARSET
        NotesFont.Color = clWindowText
        NotesFont.Height = -11
        NotesFont.Name = 'Tahoma'
        NotesFont.Style = []
        ParentFont = False
        Picture.Data = {
          424D360300000000000036000000280000001000000010000000010018000000
          000000030000C40E0000C40E00000000000000000000FFFFFFFFFFFFFFFFFFFF
          FFFFB9A897867565755545654536553D27654527653D27553D27553D27553D27
          553D27553D27FFFFFFFFFFFFFFFFFFFFFFFFB9A897FFEEEEDCD4CBDCD4CBDCCB
          CBDCCBB9DCC2B9DCB9A8CBB1A8CBB197CBA886553D27FFFFFFFFFFFFFFFFFFFF
          FFFFB9A897FFF7EEFFF7EEFFEEDCFFE5DCFFE5CBFFDCCBFFD4B9FFD4B9FFD4B9
          CBA897553D27FFFFFFFFFFFFFFFFFFFFFFFFB9A897FFFFFFFFF7EEFFF7EEFFEE
          DCFFE5DCFFE5CBFFD4B9FFD4B9FFD4B9CBA897553D27FFFFFFFFFFFFFFFFFFFF
          FFFFB9A897FFFFFFCBA897CBA897CBA897CBA897B99F97B99F86B99786B99786
          CBB1A8754D36FFFFFFFFFFFFFFFFFFFFFFFFCBA897FFFFFFFFFFFFFFFFFFFFF7
          FFFFF7FFFFF7EEFFF7EEFFEEEEFFEEEEDCC2B9865D45FFFFFFFFFFFFFFFFFFFF
          FFFFCBB197FFFFFF977E65977E65977E65977E65977565866D65866D65866D55
          DCC2B9755545FFFFFFFFFFFFFFFFFFFFFFFFCBB197FFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFF7EEFFF7EEFFEEDCDCC2B9755D45FFFFFFFFFFFF0B65860B
          55650B3D550B3D45FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF7EEFFF7EE
          FFEEEE755D45FFFFFFFFFFFF197E9745D4EE19B1CB0B2027DC7E55DC7E55DC75
          45CB7545CB6D36CB6536CB6536CB5D27CB5527CB55270B8EB9198EA81986A855
          E5EE36D4EE0B45550B27360B3645FFE5DCFFE5DCFFE5DCFFDCCBEE9765EE8E65
          EE8655CB552736A8CB97EEFF86E5FF75EEFF55E5FF36D4EE27B9CB0B3D45EEC2
          A8EEC2A8EEB997EEB197EEB186EEA886EE8E65CB552755B1CBCBF7FFCBF7FFA8
          F7FF75EEFF55DCEE36D4EE195D75DC7E55DC7E55DC7545DC7545CB7545CB6D45
          CB6D45CB6D3645B1DC65B9DC36B9CBB9F7FF75EEFF2797A80B75860B6D86FFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF55B9DCCB
          F7FFA8F7FF279FB9FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFF55B9DC55B1DC36A8CB199FCBFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF}
        TabOrder = 2
        OnClick = ButtonAddKarObjClick
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
        Appearance.ColorMirrorDisabled = 11974326
        Appearance.ColorMirrorDisabledTo = 15921906
      end
      object ButtonDelKarObj: TAdvGlowButton
        Left = 122
        Top = 255
        Width = 100
        Height = 33
        Caption = #1059#1076#1072#1083#1080#1090#1100
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        NotesFont.Charset = DEFAULT_CHARSET
        NotesFont.Color = clWindowText
        NotesFont.Height = -11
        NotesFont.Name = 'Tahoma'
        NotesFont.Style = []
        ParentFont = False
        Picture.Data = {
          424D360300000000000036000000280000001000000010000000010018000000
          000000030000C40E0000C40E00000000000000000000FFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFAAAAAA515551959595FFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF676B670B
          0B0B414141B6B5B6FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF76716DA8A8A8
          FFFFFFFFFFFFFFFFFFFFFFFF95959519120B272019797470FFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFBBBAB92C332CB7B7B7FFFFFFFFFFFFFFFFFFFFFFFFCBCACA5A
          5A5A221B223A3A3AB3B3B3FFFFFFFFFFFFFFFFFFC7C6C7444444656565D6D6D6
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFBFBEBF3838381919194F4F4FCAC9CAFFFF
          FFC9C9C9666666303030BFBEBDFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFC1C1C1443E382720195A555AB2B0AE4F4F4F27272787837FFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFC1C2C1443E382727272727
          27362E273B343BD9D9D9FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFB5B5B5272727362E363C3C3CDBDADAFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFC5C5C5545454362E273131
          313131315F5954DCDBDBFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFF989898363636363636808080C6C4C6817C81363636777777DEDCDCFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFBDBDBD6B6B6B363636363636767676E2E2E2FFFF
          FFE1E0E076706B363636797D79E0E0E0FFFFFFFFFFFFFFFFFFA5A5A56B6B6B36
          2E27363636969296FFFFFFFFFFFFFFFFFFFFFFFFE0E0DFBDBDBD4E464E767676
          FFFFFFFFFFFFFFFFFFA5A8A53636363F3F3FB6B4B6FFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFD7D7D7999D99FFFFFFFFFFFFFFFFFFFFFFFF6C666CC2
          C0C2FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF}
        TabOrder = 3
        OnClick = ButtonDelKarObjClick
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
        Appearance.ColorMirrorDisabled = 11974326
        Appearance.ColorMirrorDisabledTo = 15921906
      end
      object ListViewKarObjOtech: TcxListView
        Left = 19
        Top = 142
        Width = 584
        Height = 67
        Columns = <
          item
            Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077
            Width = 200
          end
          item
            Caption = #1057#1086#1089#1090#1086#1103#1085#1080#1077
            Width = 250
          end
          item
            Caption = #1057#1074#1077#1076#1077#1085#1080#1103' '#1086#1073' '#1086#1095#1072#1075#1077
          end
          item
            Caption = #1050#1086#1083#1080#1095#1077#1089#1090#1074#1086' '#1079#1072#1088#1072#1078#1105#1085#1085#1086#1081' '#1087#1088#1086#1076#1091#1082#1094#1080#1080
          end
          item
            Caption = #1045#1076#1080#1085#1080#1094#1099' '#1080#1079#1084#1077#1088#1077#1085#1080#1103
          end
          item
            Caption = #1055#1088#1080#1084#1077#1095#1072#1085#1080#1077
          end
          item
            Caption = #1050#1086#1076' '#1086#1073#1098#1077#1082#1090#1072
          end>
        ShowWorkAreas = True
        Style.BorderStyle = cbsOffice11
        Style.LookAndFeel.Kind = lfOffice11
        StyleDisabled.LookAndFeel.Kind = lfOffice11
        StyleFocused.LookAndFeel.Kind = lfOffice11
        StyleHot.LookAndFeel.Kind = lfOffice11
        TabOrder = 4
        ViewStyle = vsReport
        Visible = False
      end
    end
    object TabSheet6: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'6'
      ImageIndex = 5
      object cxLabel19: TcxLabel
        Left = 20
        Top = 20
        Caption = #1053#1077#1082#1072#1088#1072#1085#1090#1080#1085#1085#1099#1077' '#1086#1073#1098#1077#1082#1090#1099', '#1086#1090#1089#1091#1090#1089#1090#1074#1091#1102#1097#1080#1077' '#1074' '#1056#1060
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object cxRadioGroup4: TcxRadioGroup
        Left = 11
        Top = 37
        Properties.Items = <
          item
            Caption = #1085#1077' '#1086#1073#1085#1072#1088#1091#1078#1077#1085#1099
          end
          item
            Caption = #1086#1073#1085#1072#1088#1091#1078#1077#1085#1099
          end>
        ItemIndex = 0
        Style.BorderColor = clNone
        Style.BorderStyle = ebsNone
        TabOrder = 1
        Height = 75
        Width = 145
      end
    end
    object TabSheet7: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'7'
      ImageIndex = 6
      object cxLabel20: TcxLabel
        Left = 20
        Top = 20
        Caption = #1054#1073#1085#1072#1088#1091#1078#1077#1085#1085#1099#1077' '#1085#1077#1082#1072#1088#1072#1085#1090#1080#1085#1085#1099#1077' '#1086#1073#1098#1077#1082#1090#1099', '#1086#1090#1089#1091#1090#1089#1090#1074#1091#1102#1097#1080#1077' '#1074' '#1056#1060
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object ListViewNekarObjOtsRF: TcxListView
        Left = 19
        Top = 54
        Width = 584
        Height = 67
        Columns = <
          item
            Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077
            Width = 200
          end
          item
            Caption = #1057#1086#1089#1090#1086#1103#1085#1080#1077
            Width = 250
          end
          item
            Caption = #1055#1088#1080#1084#1077#1095#1072#1085#1080#1077
          end
          item
            Caption = #1050#1086#1076' '#1086#1073#1098#1077#1082#1090#1072
            Width = 0
          end>
        ShowWorkAreas = True
        Style.BorderStyle = cbsOffice11
        Style.LookAndFeel.Kind = lfOffice11
        StyleDisabled.LookAndFeel.Kind = lfOffice11
        StyleFocused.LookAndFeel.Kind = lfOffice11
        StyleHot.LookAndFeel.Kind = lfOffice11
        TabOrder = 1
        ViewStyle = vsReport
      end
      object AdvGlowButton1: TAdvGlowButton
        Left = 20
        Top = 255
        Width = 102
        Height = 33
        Caption = #1044#1086#1073#1072#1074#1080#1090#1100
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        NotesFont.Charset = DEFAULT_CHARSET
        NotesFont.Color = clWindowText
        NotesFont.Height = -11
        NotesFont.Name = 'Tahoma'
        NotesFont.Style = []
        ParentFont = False
        Picture.Data = {
          424D360300000000000036000000280000001000000010000000010018000000
          000000030000C40E0000C40E00000000000000000000FFFFFFFFFFFFFFFFFFFF
          FFFFB9A897867565755545654536553D27654527653D27553D27553D27553D27
          553D27553D27FFFFFFFFFFFFFFFFFFFFFFFFB9A897FFEEEEDCD4CBDCD4CBDCCB
          CBDCCBB9DCC2B9DCB9A8CBB1A8CBB197CBA886553D27FFFFFFFFFFFFFFFFFFFF
          FFFFB9A897FFF7EEFFF7EEFFEEDCFFE5DCFFE5CBFFDCCBFFD4B9FFD4B9FFD4B9
          CBA897553D27FFFFFFFFFFFFFFFFFFFFFFFFB9A897FFFFFFFFF7EEFFF7EEFFEE
          DCFFE5DCFFE5CBFFD4B9FFD4B9FFD4B9CBA897553D27FFFFFFFFFFFFFFFFFFFF
          FFFFB9A897FFFFFFCBA897CBA897CBA897CBA897B99F97B99F86B99786B99786
          CBB1A8754D36FFFFFFFFFFFFFFFFFFFFFFFFCBA897FFFFFFFFFFFFFFFFFFFFF7
          FFFFF7FFFFF7EEFFF7EEFFEEEEFFEEEEDCC2B9865D45FFFFFFFFFFFFFFFFFFFF
          FFFFCBB197FFFFFF977E65977E65977E65977E65977565866D65866D65866D55
          DCC2B9755545FFFFFFFFFFFFFFFFFFFFFFFFCBB197FFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFF7EEFFF7EEFFEEDCDCC2B9755D45FFFFFFFFFFFF0B65860B
          55650B3D550B3D45FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF7EEFFF7EE
          FFEEEE755D45FFFFFFFFFFFF197E9745D4EE19B1CB0B2027DC7E55DC7E55DC75
          45CB7545CB6D36CB6536CB6536CB5D27CB5527CB55270B8EB9198EA81986A855
          E5EE36D4EE0B45550B27360B3645FFE5DCFFE5DCFFE5DCFFDCCBEE9765EE8E65
          EE8655CB552736A8CB97EEFF86E5FF75EEFF55E5FF36D4EE27B9CB0B3D45EEC2
          A8EEC2A8EEB997EEB197EEB186EEA886EE8E65CB552755B1CBCBF7FFCBF7FFA8
          F7FF75EEFF55DCEE36D4EE195D75DC7E55DC7E55DC7545DC7545CB7545CB6D45
          CB6D45CB6D3645B1DC65B9DC36B9CBB9F7FF75EEFF2797A80B75860B6D86FFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF55B9DCCB
          F7FFA8F7FF279FB9FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFF55B9DC55B1DC36A8CB199FCBFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF}
        TabOrder = 2
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
        Appearance.ColorMirrorDisabled = 11974326
        Appearance.ColorMirrorDisabledTo = 15921906
      end
      object AdvGlowButton2: TAdvGlowButton
        Left = 122
        Top = 255
        Width = 100
        Height = 33
        Caption = #1059#1076#1072#1083#1080#1090#1100
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        NotesFont.Charset = DEFAULT_CHARSET
        NotesFont.Color = clWindowText
        NotesFont.Height = -11
        NotesFont.Name = 'Tahoma'
        NotesFont.Style = []
        ParentFont = False
        Picture.Data = {
          424D360300000000000036000000280000001000000010000000010018000000
          000000030000C40E0000C40E00000000000000000000FFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFAAAAAA515551959595FFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF676B670B
          0B0B414141B6B5B6FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF76716DA8A8A8
          FFFFFFFFFFFFFFFFFFFFFFFF95959519120B272019797470FFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFBBBAB92C332CB7B7B7FFFFFFFFFFFFFFFFFFFFFFFFCBCACA5A
          5A5A221B223A3A3AB3B3B3FFFFFFFFFFFFFFFFFFC7C6C7444444656565D6D6D6
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFBFBEBF3838381919194F4F4FCAC9CAFFFF
          FFC9C9C9666666303030BFBEBDFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFC1C1C1443E382720195A555AB2B0AE4F4F4F27272787837FFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFC1C2C1443E382727272727
          27362E273B343BD9D9D9FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFB5B5B5272727362E363C3C3CDBDADAFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFC5C5C5545454362E273131
          313131315F5954DCDBDBFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFF989898363636363636808080C6C4C6817C81363636777777DEDCDCFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFBDBDBD6B6B6B363636363636767676E2E2E2FFFF
          FFE1E0E076706B363636797D79E0E0E0FFFFFFFFFFFFFFFFFFA5A5A56B6B6B36
          2E27363636969296FFFFFFFFFFFFFFFFFFFFFFFFE0E0DFBDBDBD4E464E767676
          FFFFFFFFFFFFFFFFFFA5A8A53636363F3F3FB6B4B6FFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFD7D7D7999D99FFFFFFFFFFFFFFFFFFFFFFFF6C666CC2
          C0C2FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF}
        TabOrder = 3
        OnClick = AdvGlowButton2Click
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
        Appearance.ColorMirrorDisabled = 11974326
        Appearance.ColorMirrorDisabledTo = 15921906
      end
      object ListViewNekarObjOtsRFOtech: TcxListView
        Left = 19
        Top = 142
        Width = 584
        Height = 67
        Columns = <
          item
            Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077
            Width = 200
          end
          item
            Caption = #1057#1086#1089#1090#1086#1103#1085#1080#1077
            Width = 250
          end
          item
            Caption = #1057#1074#1077#1076#1077#1085#1080#1103' '#1086#1073' '#1086#1095#1072#1075#1077
          end
          item
            Caption = #1050#1086#1083#1080#1095#1077#1089#1090#1074#1086' '#1079#1072#1088#1072#1078#1105#1085#1085#1086#1081' '#1087#1088#1086#1076#1091#1082#1094#1080#1080
          end
          item
            Caption = #1045#1076#1080#1085#1080#1094#1099' '#1080#1079#1084#1077#1088#1077#1085#1080#1103
          end
          item
            Caption = #1055#1088#1080#1084#1077#1095#1072#1085#1080#1077
          end
          item
            Caption = #1050#1086#1076' '#1086#1073#1098#1077#1082#1090#1072
            Width = 0
          end>
        ShowWorkAreas = True
        Style.BorderStyle = cbsOffice11
        Style.LookAndFeel.Kind = lfOffice11
        StyleDisabled.LookAndFeel.Kind = lfOffice11
        StyleFocused.LookAndFeel.Kind = lfOffice11
        StyleHot.LookAndFeel.Kind = lfOffice11
        TabOrder = 4
        ViewStyle = vsReport
        Visible = False
      end
    end
    object TabSheet8: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'8'
      ImageIndex = 7
      object cxLabel23: TcxLabel
        Left = 20
        Top = 20
        Caption = #1053#1077#1082#1072#1088#1072#1085#1090#1080#1085#1085#1099#1077' '#1086#1073#1098#1077#1082#1090#1099', '#1087#1088#1086#1095#1080#1077
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object cxRadioGroup1: TcxRadioGroup
        Left = 11
        Top = 37
        Properties.Items = <
          item
            Caption = #1085#1077' '#1086#1073#1085#1072#1088#1091#1078#1077#1085#1099
          end
          item
            Caption = #1086#1073#1085#1072#1088#1091#1078#1077#1085#1099
          end>
        ItemIndex = 0
        Style.BorderColor = clNone
        Style.BorderStyle = ebsNone
        TabOrder = 1
        Height = 75
        Width = 145
      end
    end
    object TabSheet9: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'9'
      ImageIndex = 8
      object cxLabel24: TcxLabel
        Left = 20
        Top = 20
        Caption = #1054#1073#1085#1072#1088#1091#1078#1077#1085#1085#1099#1077' '#1085#1077#1082#1072#1088#1072#1085#1090#1080#1085#1085#1099#1077' '#1086#1073#1098#1077#1082#1090#1099', '#1087#1088#1086#1095#1080#1077
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object ListViewNekarObjProchii: TcxListView
        Left = 19
        Top = 54
        Width = 584
        Height = 67
        Columns = <
          item
            Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077
            Width = 200
          end
          item
            Caption = #1057#1086#1089#1090#1086#1103#1085#1080#1077
            Width = 250
          end
          item
            Caption = #1055#1088#1080#1084#1077#1095#1072#1085#1080#1077
          end
          item
            Caption = #1050#1086#1076' '#1086#1073#1098#1077#1082#1090#1072
            Width = 0
          end>
        ShowWorkAreas = True
        Style.BorderStyle = cbsOffice11
        Style.LookAndFeel.Kind = lfOffice11
        StyleDisabled.LookAndFeel.Kind = lfOffice11
        StyleFocused.LookAndFeel.Kind = lfOffice11
        StyleHot.LookAndFeel.Kind = lfOffice11
        TabOrder = 1
        ViewStyle = vsReport
      end
      object AdvGlowButton3: TAdvGlowButton
        Left = 20
        Top = 255
        Width = 102
        Height = 33
        Caption = #1044#1086#1073#1072#1074#1080#1090#1100
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        NotesFont.Charset = DEFAULT_CHARSET
        NotesFont.Color = clWindowText
        NotesFont.Height = -11
        NotesFont.Name = 'Tahoma'
        NotesFont.Style = []
        ParentFont = False
        Picture.Data = {
          424D360300000000000036000000280000001000000010000000010018000000
          000000030000C40E0000C40E00000000000000000000FFFFFFFFFFFFFFFFFFFF
          FFFFB9A897867565755545654536553D27654527653D27553D27553D27553D27
          553D27553D27FFFFFFFFFFFFFFFFFFFFFFFFB9A897FFEEEEDCD4CBDCD4CBDCCB
          CBDCCBB9DCC2B9DCB9A8CBB1A8CBB197CBA886553D27FFFFFFFFFFFFFFFFFFFF
          FFFFB9A897FFF7EEFFF7EEFFEEDCFFE5DCFFE5CBFFDCCBFFD4B9FFD4B9FFD4B9
          CBA897553D27FFFFFFFFFFFFFFFFFFFFFFFFB9A897FFFFFFFFF7EEFFF7EEFFEE
          DCFFE5DCFFE5CBFFD4B9FFD4B9FFD4B9CBA897553D27FFFFFFFFFFFFFFFFFFFF
          FFFFB9A897FFFFFFCBA897CBA897CBA897CBA897B99F97B99F86B99786B99786
          CBB1A8754D36FFFFFFFFFFFFFFFFFFFFFFFFCBA897FFFFFFFFFFFFFFFFFFFFF7
          FFFFF7FFFFF7EEFFF7EEFFEEEEFFEEEEDCC2B9865D45FFFFFFFFFFFFFFFFFFFF
          FFFFCBB197FFFFFF977E65977E65977E65977E65977565866D65866D65866D55
          DCC2B9755545FFFFFFFFFFFFFFFFFFFFFFFFCBB197FFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFF7EEFFF7EEFFEEDCDCC2B9755D45FFFFFFFFFFFF0B65860B
          55650B3D550B3D45FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF7EEFFF7EE
          FFEEEE755D45FFFFFFFFFFFF197E9745D4EE19B1CB0B2027DC7E55DC7E55DC75
          45CB7545CB6D36CB6536CB6536CB5D27CB5527CB55270B8EB9198EA81986A855
          E5EE36D4EE0B45550B27360B3645FFE5DCFFE5DCFFE5DCFFDCCBEE9765EE8E65
          EE8655CB552736A8CB97EEFF86E5FF75EEFF55E5FF36D4EE27B9CB0B3D45EEC2
          A8EEC2A8EEB997EEB197EEB186EEA886EE8E65CB552755B1CBCBF7FFCBF7FFA8
          F7FF75EEFF55DCEE36D4EE195D75DC7E55DC7E55DC7545DC7545CB7545CB6D45
          CB6D45CB6D3645B1DC65B9DC36B9CBB9F7FF75EEFF2797A80B75860B6D86FFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF55B9DCCB
          F7FFA8F7FF279FB9FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFF55B9DC55B1DC36A8CB199FCBFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF}
        TabOrder = 2
        OnClick = AdvGlowButton3Click
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
        Appearance.ColorMirrorDisabled = 11974326
        Appearance.ColorMirrorDisabledTo = 15921906
      end
      object AdvGlowButton4: TAdvGlowButton
        Left = 122
        Top = 255
        Width = 100
        Height = 33
        Caption = #1059#1076#1072#1083#1080#1090#1100
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        NotesFont.Charset = DEFAULT_CHARSET
        NotesFont.Color = clWindowText
        NotesFont.Height = -11
        NotesFont.Name = 'Tahoma'
        NotesFont.Style = []
        ParentFont = False
        Picture.Data = {
          424D360300000000000036000000280000001000000010000000010018000000
          000000030000C40E0000C40E00000000000000000000FFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFAAAAAA515551959595FFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF676B670B
          0B0B414141B6B5B6FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF76716DA8A8A8
          FFFFFFFFFFFFFFFFFFFFFFFF95959519120B272019797470FFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFBBBAB92C332CB7B7B7FFFFFFFFFFFFFFFFFFFFFFFFCBCACA5A
          5A5A221B223A3A3AB3B3B3FFFFFFFFFFFFFFFFFFC7C6C7444444656565D6D6D6
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFBFBEBF3838381919194F4F4FCAC9CAFFFF
          FFC9C9C9666666303030BFBEBDFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFC1C1C1443E382720195A555AB2B0AE4F4F4F27272787837FFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFC1C2C1443E382727272727
          27362E273B343BD9D9D9FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFB5B5B5272727362E363C3C3CDBDADAFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFC5C5C5545454362E273131
          313131315F5954DCDBDBFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFF989898363636363636808080C6C4C6817C81363636777777DEDCDCFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFBDBDBD6B6B6B363636363636767676E2E2E2FFFF
          FFE1E0E076706B363636797D79E0E0E0FFFFFFFFFFFFFFFFFFA5A5A56B6B6B36
          2E27363636969296FFFFFFFFFFFFFFFFFFFFFFFFE0E0DFBDBDBD4E464E767676
          FFFFFFFFFFFFFFFFFFA5A8A53636363F3F3FB6B4B6FFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFD7D7D7999D99FFFFFFFFFFFFFFFFFFFFFFFF6C666CC2
          C0C2FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF}
        TabOrder = 3
        OnClick = AdvGlowButton4Click
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
        Appearance.ColorMirrorDisabled = 11974326
        Appearance.ColorMirrorDisabledTo = 15921906
      end
      object ListViewNekarObjProchiiOtech: TcxListView
        Left = 19
        Top = 142
        Width = 584
        Height = 67
        Columns = <
          item
            Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077
            Width = 200
          end
          item
            Caption = #1057#1086#1089#1090#1086#1103#1085#1080#1077
            Width = 250
          end
          item
            Caption = #1057#1074#1077#1076#1077#1085#1080#1103' '#1086#1073' '#1086#1095#1072#1075#1077
          end
          item
            Caption = #1050#1086#1083#1080#1095#1077#1089#1090#1074#1086' '#1079#1072#1088#1072#1078#1105#1085#1085#1086#1081' '#1087#1088#1086#1076#1091#1082#1094#1080#1080
          end
          item
            Caption = #1045#1076#1080#1085#1080#1094#1099' '#1080#1079#1084#1077#1088#1077#1085#1080#1103
          end
          item
            Caption = #1055#1088#1080#1084#1077#1095#1072#1085#1080#1077
          end
          item
            Caption = #1050#1086#1076' '#1086#1073#1098#1077#1082#1090#1072
            Width = 0
          end>
        ShowWorkAreas = True
        Style.BorderStyle = cbsOffice11
        Style.LookAndFeel.Kind = lfOffice11
        StyleDisabled.LookAndFeel.Kind = lfOffice11
        StyleFocused.LookAndFeel.Kind = lfOffice11
        StyleHot.LookAndFeel.Kind = lfOffice11
        TabOrder = 4
        ViewStyle = vsReport
        Visible = False
      end
    end
    object TabSheet10: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'10'
      ImageIndex = 10
      object LabelInfo: TcxLabel
        Left = 18
        Top = 20
        AutoSize = False
        Caption = #1069#1082#1089#1087#1077#1088#1090#1080#1079#1072' '#1086#1082#1086#1085#1095#1077#1085#1072'!'
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
        Properties.WordWrap = True
        Height = 30
        Width = 223
      end
      object cxLabel26: TcxLabel
        Left = 18
        Top = 61
        AutoSize = False
        Caption = #1047#1072#1087#1080#1089#1080' '#1087#1086' '#1101#1082#1089#1087#1077#1088#1090#1080#1079#1077' '#1091#1089#1087#1077#1096#1085#1086' '#1079#1072#1085#1077#1089#1077#1085#1099' '#1074' '#1073#1072#1079#1091' '#1076#1072#1085#1085#1099#1093'!'
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = []
        Style.IsFontAssigned = True
        Properties.WordWrap = True
        Height = 36
        Width = 463
      end
    end
  end
  object ButtonForward: TAdvGlowButton
    Left = 378
    Top = 369
    Width = 100
    Height = 33
    Caption = #1055#1088#1086#1076#1086#1083#1078#1080#1090#1100' >'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    NotesFont.Charset = DEFAULT_CHARSET
    NotesFont.Color = clWindowText
    NotesFont.Height = -11
    NotesFont.Name = 'Tahoma'
    NotesFont.Style = []
    ParentFont = False
    TabOrder = 1
    OnClick = ButtonForwardClick
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
    Appearance.ColorMirrorDisabled = 11974326
    Appearance.ColorMirrorDisabledTo = 15921906
  end
  object ButtonCancel: TAdvGlowButton
    Left = 522
    Top = 369
    Width = 100
    Height = 33
    Caption = #1054#1090#1084#1077#1085#1072
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    NotesFont.Charset = DEFAULT_CHARSET
    NotesFont.Color = clWindowText
    NotesFont.Height = -11
    NotesFont.Name = 'Tahoma'
    NotesFont.Style = []
    ParentFont = False
    TabOrder = 2
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
    Appearance.ColorMirrorDisabled = 11974326
    Appearance.ColorMirrorDisabledTo = 15921906
  end
  object ButtonBackward: TAdvGlowButton
    Left = 276
    Top = 369
    Width = 102
    Height = 33
    Caption = '< '#1053#1072#1079#1072#1076
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    NotesFont.Charset = DEFAULT_CHARSET
    NotesFont.Color = clWindowText
    NotesFont.Height = -11
    NotesFont.Name = 'Tahoma'
    NotesFont.Style = []
    ParentFont = False
    TabOrder = 3
    OnClick = ButtonBackwardClick
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
    Appearance.ColorMirrorDisabled = 11974326
    Appearance.ColorMirrorDisabledTo = 15921906
  end
  object AdvPanel1: TAdvPanel
    Left = 0
    Top = 0
    Width = 636
    Height = 33
    Align = alTop
    BevelOuter = bvNone
    Color = clCaptionText
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 4
    UseDockManager = True
    Version = '1.7.8.1'
    Caption.Color = clHighlight
    Caption.ColorTo = clNone
    Caption.Font.Charset = DEFAULT_CHARSET
    Caption.Font.Color = clHighlightText
    Caption.Font.Height = -11
    Caption.Font.Name = 'Tahoma'
    Caption.Font.Style = []
    Caption.Text = #1052#1072#1089#1090#1077#1088' '#1088#1077#1075#1080#1089#1090#1088#1072#1094#1080#1080
    ColorTo = clActiveBorder
    Hover = True
    HoverColor = clMenu
    HoverFontColor = clGradientInactiveCaption
    StatusBar.BorderColor = clGray
    StatusBar.BorderStyle = bsSingle
    StatusBar.Font.Charset = DEFAULT_CHARSET
    StatusBar.Font.Color = clWindowText
    StatusBar.Font.Height = -11
    StatusBar.Font.Name = 'Tahoma'
    StatusBar.Font.Style = []
    StatusBar.ColorTo = clScrollBar
    FullHeight = 0
    object LabelPodskazka: TcxLabel
      Left = 7
      Top = 3
      Caption = #1055#1086#1076#1089#1082#1072#1079#1082#1072
      ParentFont = False
      Style.Font.Charset = DEFAULT_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -11
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = [fsBold]
      Style.LookAndFeel.Kind = lfOffice11
      Style.IsFontAssigned = True
      StyleDisabled.LookAndFeel.Kind = lfOffice11
      StyleFocused.LookAndFeel.Kind = lfOffice11
      StyleHot.LookAndFeel.Kind = lfOffice11
      Transparent = True
    end
  end
  object ADODataSet1: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from SPETSIALISTS_LAB ORDER BY NAMESPETSIALIST;'
    CommandTimeout = 600
    Parameters = <>
    Left = 5
    Top = 378
  end
  object DataSource1: TDataSource
    DataSet = ADODataSet1
    Left = 33
    Top = 378
  end
  object ADODataSet2: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from KAROBJECTS ORDER BY TYPEOBJ, NAMEOBJ;'
    CommandTimeout = 600
    Parameters = <>
    Left = 68
    Top = 378
  end
  object DataSource2: TDataSource
    DataSet = ADODataSet2
    Left = 96
    Top = 378
  end
  object ADODataSet3: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from TRANSPORT;'
    CommandTimeout = 600
    Parameters = <>
    Left = 132
    Top = 378
  end
  object DataSource3: TDataSource
    DataSet = ADODataSet3
    Left = 160
    Top = 378
  end
  object ADODataSet4: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from NEKAROBJECTS_NEZARRF;'
    CommandTimeout = 600
    Parameters = <>
    Left = 196
    Top = 378
  end
  object DataSource4: TDataSource
    DataSet = ADODataSet4
    Left = 224
    Top = 378
  end
  object ADOCommand1: TADOCommand
    Connection = FrmMainWindow.ADOConnection1
    Parameters = <>
    Left = 7
    Top = 324
  end
  object ADODataSet5: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandTimeout = 600
    Parameters = <>
    Left = 35
    Top = 324
  end
end
