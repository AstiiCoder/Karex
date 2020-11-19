object FrmDialog1: TFrmDialog1
  Left = 475
  Top = 307
  BorderStyle = bsDialog
  Caption = #1052#1072#1089#1090#1077#1088' '#1088#1077#1075#1080#1089#1090#1088#1072#1094#1080#1080' '#1086#1073#1088#1072#1079#1094#1086#1074
  ClientHeight = 415
  ClientWidth = 635
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object DialogWizard: TPageControl
    Left = 0
    Top = 33
    Width = 635
    Height = 329
    ActivePage = TabSheet1
    Align = alTop
    Style = tsButtons
    TabOrder = 0
    object TabSheet1: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'1'
      object LabelNum: TcxLabel
        Left = 152
        Top = 176
        Caption = #1053#1086#1084#1077#1088' '#1087#1088#1086#1090#1086#1082#1086#1083#1072' '#1082#1072#1088#1072#1085#1090#1080#1085#1085#1086#1081' '#1101#1082#1089#1087#1077#1088#1090#1080#1079#1099
      end
      object LabelDlgName: TcxLabel
        Left = 20
        Top = 20
        Caption = #1056#1077#1075#1080#1089#1090#1088#1072#1094#1080#1103' '#1086#1073#1088#1072#1079#1094#1086#1074
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object LabelDate: TcxLabel
        Left = 20
        Top = 176
        Caption = #1044#1072#1090#1072' '#1088#1077#1075#1080#1089#1090#1088#1072#1094#1080#1080
      end
      object DateEdit1: TcxDateEdit
        Left = 20
        Top = 212
        EditValue = 0d
        Properties.DateButtons = []
        Properties.YearsInMonthList = False
        Properties.OnEditValueChanged = DateEdit1PropertiesEditValueChanged
        TabOrder = 0
        Width = 98
      end
      object CalcEdit1: TcxCalcEdit
        Left = 152
        Top = 212
        EditValue = 0
        Enabled = False
        Properties.ImmediatePost = True
        Properties.ReadOnly = False
        Style.ButtonStyle = btsOffice11
        TabOrder = 1
        Width = 97
      end
      object RadioGroupImport: TcxRadioGroup
        Left = 20
        Top = 51
        Properties.Items = <
          item
            Caption = #1080#1084#1087#1086#1088#1090#1085#1099#1081
          end
          item
            Caption = #1086#1090#1077#1095#1077#1089#1090#1074#1077#1085#1085#1099#1081
          end>
        Properties.OnEditValueChanged = RadioGroupImportPropertiesEditValueChanged
        ItemIndex = 0
        Style.BorderStyle = ebsOffice11
        TabOrder = 2
        Height = 94
        Width = 116
      end
      object RadioGroupNumType: TcxRadioGroup
        Left = 283
        Top = 206
        Caption = #1085#1086#1084#1077#1088#1072#1094#1080#1103' '#1087#1088#1086#1090#1086#1082#1086#1083#1086#1074
        Properties.Items = <
          item
            Caption = #1072#1074#1090#1086#1084#1072#1090#1080#1095#1077#1089#1082#1072#1103
          end
          item
            Caption = #1079#1072#1076#1072#1090#1100' '#1085#1086#1084#1077#1088
          end>
        Properties.OnChange = RadioGroupNumTypePropertiesChange
        ItemIndex = 0
        TabOrder = 6
        Height = 91
        Width = 185
      end
    end
    object TabSheet2: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'2'
      ImageIndex = 1
      object cxLabel1: TcxLabel
        Left = 20
        Top = 20
        Caption = 
          #1057#1086#1087#1088#1086#1074#1086#1076#1080#1090#1077#1083#1100#1085#1099#1081' '#1076#1086#1082#1091#1084#1077#1085#1090' '#1056#1086#1089#1089#1077#1083#1100#1093#1086#1079#1085#1072#1076#1079#1086#1088#1072', '#1083#1080#1073#1086' '#1056#1077#1092#1077#1088#1077#1085#1090#1085#1086#1075#1086' '#1094 +
          #1077#1085#1090#1088#1072
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object LabelNameUpr: TcxLabel
        Left = 18
        Top = 98
        Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077' '#1086#1088#1075#1072#1085#1080#1079#1072#1094#1080#1080', '#1074#1099#1076#1072#1074#1096#1077#1081' '#1076#1086#1082#1091#1084#1077#1085#1090
      end
      object ButtonNameUpr: TcxButton
        Left = 583
        Top = 98
        Width = 20
        Height = 19
        Caption = '...'
        TabOrder = 2
        OnClick = ButtonNameUprClick
        OnExit = ButtonNameUprExit
      end
      object LabelNumEtiketki: TcxLabel
        Left = 18
        Top = 213
        Caption = #1053#1086#1084#1077#1088' '#1076#1086#1082#1091#1084#1077#1085#1090#1072
      end
      object LabelDate3: TcxLabel
        Left = 18
        Top = 252
        Caption = #1044#1072#1090#1072' '#1074#1099#1076#1072#1095#1080' '#1076#1086#1082#1091#1084#1077#1085#1090#1072
      end
      object DateEdit3: TcxDateEdit
        Left = 277
        Top = 251
        Properties.DateButtons = []
        Properties.ImmediatePost = True
        Properties.YearsInMonthList = False
        Properties.OnEditValueChanged = DateEdit3PropertiesEditValueChanged
        TabOrder = 8
        Width = 98
      end
      object LabelPunktName: TcxLabel
        Left = 17
        Top = 128
        AutoSize = False
        Caption = 
          #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077' '#1087#1091#1085#1082#1090#1072' '#1087#1086' '#1082#1072#1088#1072#1085#1090#1080#1085#1091' '#1088#1072#1089#1090#1077#1085#1080#1081', '#1087#1086#1076#1088#1072#1079#1076#1077#1083#1077#1085#1080#1103', '#1086#1090#1076#1077#1083#1072 +
          ' '#1080' '#1090'.'#1087'.  '
        Properties.WordWrap = True
        Height = 33
        Width = 256
      end
      object ButtonPunktName: TcxButton
        Left = 582
        Top = 135
        Width = 20
        Height = 19
        Caption = '...'
        TabOrder = 4
        OnClick = ButtonPunktNameClick
        OnExit = ButtonPunktNameExit
      end
      object LabelSpetsialist: TcxLabel
        Left = 17
        Top = 173
        Caption = #1060'.'#1048'.'#1054'. '#1089#1087#1077#1094#1080#1072#1083#1080#1089#1090#1072
      end
      object ButtonBoxSpetsialist: TcxButton
        Left = 582
        Top = 173
        Width = 20
        Height = 19
        Caption = '...'
        TabOrder = 6
        OnClick = ButtonBoxSpetsialistClick
        OnExit = ButtonBoxSpetsialistExit
      end
      object ComboBoxSpetsialist: TcxLookupComboBox
        Left = 276
        Top = 172
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
        EditValue = 0
        Style.BorderStyle = ebsOffice11
        Style.ButtonStyle = btsOffice11
        TabOrder = 5
        Width = 308
      end
      object TextEditNumEtiketki: TcxTextEdit
        Left = 277
        Top = 212
        Properties.OnEditValueChanged = TextEditNumEtiketkiPropertiesEditValueChanged
        TabOrder = 7
        Text = '0'
        Width = 327
      end
      object ComboBoxPunkt: TcxLookupComboBox
        Left = 276
        Top = 134
        Properties.DropDownRows = 20
        Properties.ImmediatePost = True
        Properties.KeyFieldNames = 'IDKARPUNKT'
        Properties.ListColumns = <
          item
            Caption = #1050#1086#1076' '#1087#1091#1085#1082#1090#1072
            MinWidth = 0
            Width = 0
            FieldName = 'IDKARPUNKT'
          end
          item
            Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077
            FieldName = 'NAMEKARPUNKT'
          end>
        Properties.ListFieldIndex = 1
        Properties.ListOptions.ShowHeader = False
        Properties.ListSource = DataSource2
        EditValue = 0
        Style.BorderStyle = ebsOffice11
        Style.ButtonStyle = btsOffice11
        TabOrder = 3
        Width = 308
      end
      object ComboBoxUprNadz: TcxLookupComboBox
        Left = 277
        Top = 97
        ParentShowHint = False
        Properties.DropDownRows = 20
        Properties.DropDownWidth = 600
        Properties.ImmediatePost = True
        Properties.KeyFieldNames = 'IDUPRSHNADZ'
        Properties.ListColumns = <
          item
            Caption = #1050#1086#1076' '#1091#1087#1088#1072#1074#1083#1077#1085#1080#1103
            MinWidth = 0
            Width = 0
            FieldName = 'IDUPRSHNADZ'
          end
          item
            Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077
            MinWidth = 300
            Width = 300
            FieldName = 'NAMEUPRSHNADZ'
          end>
        Properties.ListFieldIndex = 1
        Properties.ListOptions.ShowHeader = False
        Properties.ListSource = DataSource3
        Properties.OnEditValueChanged = ComboBoxUprNadzPropertiesEditValueChanged
        EditValue = 0
        ShowHint = True
        Style.BorderStyle = ebsOffice11
        Style.ButtonStyle = btsOffice11
        TabOrder = 1
        Width = 308
      end
      object CheckBoxNetEtiketki: TcxCheckBox
        Left = 494
        Top = 19
        Caption = #1086#1090#1089#1091#1090#1089#1090#1074#1091#1077#1090
        Properties.OnEditValueChanged = CheckBoxNetEtiketkiPropertiesEditValueChanged
        TabOrder = 15
        Width = 94
      end
      object LabelNazvDok: TcxLabel
        Left = 17
        Top = 60
        AutoSize = False
        Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077' '#1076#1086#1082#1091#1084#1077#1085#1090#1072
        Height = 23
        Width = 137
      end
      object ComboBoxNazvDok: TcxLookupComboBox
        Left = 276
        Top = 59
        ParentShowHint = False
        Properties.DropDownRows = 20
        Properties.DropDownWidth = 500
        Properties.ImmediatePost = True
        Properties.KeyFieldNames = 'IDDOCVIDS'
        Properties.ListColumns = <
          item
            Caption = #1050#1086#1076' '#1074#1080#1076#1072' '#1076#1086#1082#1091#1084#1077#1085#1090#1072
            MinWidth = 0
            Width = 0
            FieldName = 'IDDOCVIDS'
          end
          item
            Caption = #1042#1080#1076' '#1076#1086#1082#1091#1084#1077#1085#1090#1072
            MinWidth = 300
            FieldName = 'DOCVIDS'
          end>
        Properties.ListFieldIndex = 1
        Properties.ListOptions.ShowHeader = False
        Properties.ListSource = DataSource13
        Properties.OnEditValueChanged = ComboBoxNazvDokPropertiesEditValueChanged
        EditValue = 0
        ShowHint = True
        Style.BorderStyle = ebsOffice11
        Style.ButtonStyle = btsOffice11
        TabOrder = 17
        Width = 308
      end
      object ButtonNazvDok: TcxButton
        Left = 582
        Top = 60
        Width = 20
        Height = 19
        Caption = '...'
        TabOrder = 0
        OnClick = ButtonNazvDokClick
        OnExit = ButtonNazvDokExit
      end
    end
    object TabSheet3: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'3'
      ImageIndex = 2
      object TextEditOrgName: TcxTextEdit
        Left = 192
        Top = 53
        Enabled = False
        TabOrder = 3
        Visible = False
        Width = 30
      end
      object TextEditAdres: TcxTextEdit
        Left = 277
        Top = 133
        Enabled = False
        TabOrder = 4
        Width = 326
      end
      object cxLabel4: TcxLabel
        Left = 18
        Top = 173
        Caption = #1053#1086#1084#1077#1088' '#1076#1086#1075#1086#1074#1086#1088#1072
      end
      object LabelNameClient: TcxLabel
        Left = 18
        Top = 58
        Caption = #1040#1073#1073#1088#1077#1074#1080#1072#1090#1091#1088#1072' '#1082#1083#1080#1077#1085#1090#1072
      end
      object cxLabel5: TcxLabel
        Left = 20
        Top = 20
        Caption = #1050#1083#1080#1077#1085#1090
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object TextEditDogovorNum: TcxTextEdit
        Left = 277
        Top = 171
        Properties.OnEditValueChanged = TextEditDogovorNumPropertiesEditValueChanged
        TabOrder = 2
        Text = '0'
        Width = 97
      end
      object cxButton1: TcxButton
        Left = 583
        Top = 95
        Width = 20
        Height = 19
        Caption = '...'
        TabOrder = 1
        OnClick = cxButton1Click
        OnExit = cxButton1Exit
      end
      object LabelOrgName: TcxLabel
        Left = 17
        Top = 95
        Caption = #1053#1072#1079#1074#1072#1085#1080#1077' '#1086#1088#1075#1072#1085#1080#1079#1072#1094#1080#1080' ('#1060'.'#1048'.'#1054'. '#1095#1072#1089#1090#1085#1086#1075#1086' '#1083#1080#1094#1072') '
      end
      object LabelAdres: TcxLabel
        Left = 17
        Top = 133
        Caption = #1070#1088#1080#1076#1080#1095#1077#1089#1082#1080#1081' '#1072#1076#1088#1077#1089
      end
      object ComboBoxOrganiz: TcxLookupComboBox
        Left = 277
        Top = 94
        Properties.DropDownRows = 20
        Properties.ImmediatePost = True
        Properties.KeyFieldNames = 'IDCLIENT'
        Properties.ListColumns = <
          item
            Caption = #1050#1086#1076' '#1082#1083#1080#1077#1085#1090#1072
            MinWidth = 0
            Width = 0
            FieldName = 'IDCLIENT'
          end
          item
            Caption = #1053#1072#1079#1074#1072#1085#1080#1077
            FieldName = 'NAMECLIENT'
          end
          item
            MinWidth = 0
            Width = 0
            FieldName = 'ABBREVCLIENT'
          end
          item
            FieldName = 'ADRESSCLIENT'
          end>
        Properties.ListFieldIndex = 1
        Properties.ListOptions.ShowHeader = False
        Properties.ListSource = DataSource4
        Properties.OnChange = ComboBoxOrganizPropertiesChange
        Properties.OnEditValueChanged = ComboBoxOrganizPropertiesEditValueChanged
        EditValue = 0
        Style.BorderStyle = ebsOffice11
        Style.ButtonStyle = btsOffice11
        TabOrder = 0
        Width = 308
      end
      object TextEditAbbr: TcxTextEdit
        Left = 277
        Top = 53
        Enabled = False
        TabOrder = 10
        Width = 326
      end
    end
    object TabSheet4: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'4'
      ImageIndex = 3
      object cxLabel6: TcxLabel
        Left = 20
        Top = 20
        Caption = #1047#1072#1103#1074#1082#1072' '#1082#1083#1080#1077#1085#1090#1072
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object LabelZayavkaNum: TcxLabel
        Left = 20
        Top = 54
        Caption = #1047#1072#1103#1074#1082#1072' '#1082#1083#1080#1077#1085#1090#1072' '#8470
      end
      object LabelDateEdit4: TcxLabel
        Left = 159
        Top = 54
        Caption = #1044#1072#1090#1072' '#1079#1072#1103#1074#1082#1080
      end
      object DateEdit4: TcxDateEdit
        Left = 159
        Top = 89
        Properties.DateButtons = []
        Properties.ImmediatePost = True
        Properties.YearsInMonthList = False
        Properties.OnEditValueChanged = DateEdit4PropertiesEditValueChanged
        TabOrder = 1
        Width = 98
      end
      object TextEditZayavkaNum: TcxTextEdit
        Left = 22
        Top = 89
        Properties.OnEditValueChanged = TextEditZayavkaNumPropertiesEditValueChanged
        TabOrder = 0
        Text = '0'
        Width = 97
      end
      object CheckBoxNetZayavki: TcxCheckBox
        Left = 129
        Top = 19
        Caption = #1086#1090#1089#1091#1090#1089#1090#1074#1091#1077#1090
        Properties.OnEditValueChanged = CheckBoxNetZayavkiPropertiesEditValueChanged
        TabOrder = 5
        Width = 96
      end
    end
    object TabSheet5: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'5'
      ImageIndex = 4
      object cxLabel9: TcxLabel
        Left = 20
        Top = 78
        AutoSize = False
        Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077' '#1087#1072#1088#1090#1080#1080' '#1087#1088#1086#1076#1091#1082#1094#1080#1080',  '#1086#1090' '#1082#1086#1090#1086#1088#1086#1081' '#1086#1090#1086#1073#1088#1072#1085' '#1086#1073#1088#1072#1079#1077#1094'  '
        Properties.WordWrap = True
        Height = 35
        Width = 237
      end
      object ButtonSprPartii: TcxButton
        Left = 582
        Top = 86
        Width = 20
        Height = 19
        Caption = '...'
        TabOrder = 3
        OnClick = ButtonSprPartiiClick
        OnExit = ButtonSprPartiiExit
      end
      object cxLabel10: TcxLabel
        Left = 20
        Top = 124
        Caption = #1050#1083#1072#1089#1089#1080#1092#1080#1082#1072#1094#1080#1103' '#1087#1088#1086#1076#1091#1082#1094#1080#1080
      end
      object ButtonSprKlass: TcxButton
        Left = 582
        Top = 124
        Width = 19
        Height = 19
        Caption = '...'
        TabOrder = 5
        OnClick = ButtonSprKlassClick
        OnExit = ButtonSprKlassExit
      end
      object cxLabel12: TcxLabel
        Left = 20
        Top = 200
        Caption = #1055#1088#1086#1080#1089#1093#1086#1078#1076#1077#1085#1080#1077
      end
      object ButtonSprProishozd: TcxButton
        Left = 582
        Top = 200
        Width = 20
        Height = 19
        Caption = '...'
        TabOrder = 9
        OnClick = ButtonSprProishozdClick
        OnExit = ButtonSprProishozdExit
      end
      object cxLabel13: TcxLabel
        Left = 20
        Top = 277
        Caption = #1050#1086#1083#1080#1095#1077#1089#1090#1074#1086' '#1084#1077#1089#1090
      end
      object CalcEditEdIzm: TcxCalcEdit
        Left = 277
        Top = 276
        EditValue = 0
        Properties.ImmediatePost = True
        Properties.OnEditValueChanged = CalcEditEdIzmPropertiesEditValueChanged
        TabOrder = 13
        Width = 97
      end
      object ButtonSprVes: TcxButton
        Left = 583
        Top = 239
        Width = 20
        Height = 19
        Caption = '...'
        TabOrder = 12
        OnClick = ButtonSprVesClick
        OnExit = ButtonSprVesExit
      end
      object ButtonSprEdIzm: TcxButton
        Left = 583
        Top = 277
        Width = 20
        Height = 19
        Caption = '...'
        TabOrder = 15
        OnClick = ButtonSprEdIzmClick
        OnExit = ButtonSprEdIzmExit
      end
      object cxLabel14: TcxLabel
        Left = 20
        Top = 20
        Caption = #1057#1074#1077#1076#1077#1085#1080#1103' '#1086#1073' '#1086#1073#1088#1072#1079#1094#1077
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object cxLabel15: TcxLabel
        Left = 20
        Top = 49
        Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077' '#1086#1073#1088#1072#1079#1094#1072
      end
      object ButtonSprObrazets: TcxButton
        Left = 583
        Top = 49
        Width = 20
        Height = 19
        Caption = '...'
        TabOrder = 1
        OnClick = ButtonSprObrazetsClick
        OnExit = ButtonSprObrazetsExit
      end
      object cxLabel16: TcxLabel
        Left = 20
        Top = 239
        Caption = #1042#1077#1089
      end
      object CalcEditVes: TcxCalcEdit
        Left = 277
        Top = 238
        EditValue = 0.000000000000000000
        Properties.ImmediatePost = True
        Properties.OnEditValueChanged = CalcEditVesPropertiesEditValueChanged
        TabOrder = 10
        Width = 97
      end
      object cxLabel17: TcxLabel
        Left = 397
        Top = 240
        Caption = #1074' '#1077#1076#1080#1085#1080#1094#1072#1093' '#1080#1079#1084'.'
      end
      object cxLabel18: TcxLabel
        Left = 397
        Top = 277
        Caption = #1074' '#1077#1076#1080#1085#1080#1094#1072#1093' '#1080#1079#1084'.'
      end
      object TextEditSort: TcxTextEdit
        Left = 277
        Top = 160
        Enabled = False
        Properties.OnEditValueChanged = TextEditSortPropertiesEditValueChanged
        TabOrder = 7
        Text = #1085#1077#1090
        Width = 326
      end
      object CheckBoxSort: TcxCheckBox
        Left = 20
        Top = 162
        Caption = #1057#1086#1088#1090', '#1088#1077#1087#1088#1086#1076#1091#1082#1094#1080#1103', '#1082#1083#1072#1089#1089
        Properties.OnChange = CheckBoxSortPropertiesChange
        TabOrder = 6
        Width = 221
      end
      object ComboBoxObrazets: TcxLookupComboBox
        Left = 277
        Top = 48
        Properties.DropDownRows = 20
        Properties.ImmediatePost = True
        Properties.KeyFieldNames = 'IDOBR'
        Properties.ListColumns = <
          item
            Caption = #1050#1086#1076' '#1086#1073#1088#1072#1079#1094#1072
            MinWidth = 0
            Width = 0
            FieldName = 'IDOBR'
          end
          item
            Caption = #1053#1072#1079#1074#1072#1085#1080#1077
            FieldName = 'NAMEOBR'
          end>
        Properties.ListFieldIndex = 1
        Properties.ListOptions.ShowHeader = False
        Properties.ListSource = DataSource5
        EditValue = 0
        Style.BorderStyle = ebsOffice11
        Style.ButtonStyle = btsOffice11
        TabOrder = 0
        Width = 308
      end
      object ComboBoxPartii: TcxLookupComboBox
        Left = 276
        Top = 85
        Properties.DropDownRows = 20
        Properties.ImmediatePost = True
        Properties.KeyFieldNames = 'IDPARTII'
        Properties.ListColumns = <
          item
            Caption = #1050#1086#1076' '#1087#1072#1088#1090#1080#1080
            MinWidth = 0
            Width = 0
            FieldName = 'IDPARTII'
          end
          item
            Caption = #1053#1072#1079#1074#1072#1085#1080#1077
            FieldName = 'NAMEPARTII'
          end>
        Properties.ListFieldIndex = 1
        Properties.ListOptions.ShowHeader = False
        Properties.ListSource = DataSource6
        EditValue = 0
        Style.BorderStyle = ebsOffice11
        Style.ButtonStyle = btsOffice11
        TabOrder = 2
        Width = 308
      end
      object ComboBoxKlass: TcxLookupComboBox
        Left = 276
        Top = 123
        Properties.DropDownRows = 20
        Properties.ImmediatePost = True
        Properties.KeyFieldNames = 'IDKLASS'
        Properties.ListColumns = <
          item
            Caption = #1050#1086#1076' '#1082#1083#1072#1089#1089#1080#1092#1080#1082#1072#1094#1080#1080
            MinWidth = 0
            Width = 0
            FieldName = 'IDKLASS'
          end
          item
            Caption = #1053#1072#1079#1074#1072#1085#1080#1077
            FieldName = 'NAMEKLASS'
          end>
        Properties.ListFieldIndex = 1
        Properties.ListOptions.ShowHeader = False
        Properties.ListSource = DataSource7
        EditValue = 0
        Style.BorderStyle = ebsOffice11
        Style.ButtonStyle = btsOffice11
        TabOrder = 4
        Width = 308
      end
      object ComboBoxProishozhd: TcxLookupComboBox
        Left = 276
        Top = 199
        Properties.DropDownRows = 20
        Properties.ImmediatePost = True
        Properties.KeyFieldNames = 'IDPROISH'
        Properties.ListColumns = <
          item
            Caption = #1050#1086#1076
            MinWidth = 0
            Width = 0
            FieldName = 'IDPROISH'
          end
          item
            Caption = #1053#1072#1079#1074#1072#1085#1080#1077
            FieldName = 'NAMEPROISH'
          end>
        Properties.ListFieldIndex = 1
        Properties.ListOptions.ShowHeader = False
        Properties.ListSource = DataSource8
        EditValue = 0
        Style.BorderStyle = ebsOffice11
        Style.ButtonStyle = btsOffice11
        TabOrder = 8
        Width = 308
      end
      object ComboBoxVes: TcxLookupComboBox
        Left = 489
        Top = 238
        Properties.DropDownRows = 20
        Properties.ImmediatePost = True
        Properties.KeyFieldNames = 'IDMEASURE'
        Properties.ListColumns = <
          item
            Caption = #1050#1086#1076
            MinWidth = 0
            Width = 0
            FieldName = 'IDMEASURE'
          end
          item
            Caption = #1053#1072#1079#1074#1072#1085#1080#1077
            FieldName = 'NAMEMEASURE'
          end>
        Properties.ListFieldIndex = 1
        Properties.ListOptions.ShowHeader = False
        Properties.ListSource = DataSource9
        EditValue = 0
        Style.BorderStyle = ebsOffice11
        Style.ButtonStyle = btsOffice11
        TabOrder = 11
        Width = 96
      end
      object ComboBoxEdIzm: TcxLookupComboBox
        Left = 489
        Top = 276
        Properties.DropDownRows = 20
        Properties.ImmediatePost = True
        Properties.KeyFieldNames = 'IDEDIZM'
        Properties.ListColumns = <
          item
            Caption = #1050#1086#1076
            MinWidth = 0
            Width = 0
            FieldName = 'IDEDIZM'
          end
          item
            Caption = #1053#1072#1079#1074#1072#1085#1080#1077
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
    end
    object TabSheet6: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'6'
      ImageIndex = 5
      object cxLabel19: TcxLabel
        Left = 20
        Top = 20
        Caption = #1054#1073#1088#1072#1079#1094#1099
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object LabelKolich: TcxLabel
        Left = 20
        Top = 52
        Caption = #1050#1086#1083#1080#1095#1077#1089#1090#1074#1086' '#1086#1073#1088#1072#1079#1094#1086#1074
      end
      object CalcEditKolich: TcxCalcEdit
        Left = 20
        Top = 88
        EditValue = 0
        Properties.ImmediatePost = True
        Properties.OnEditValueChanged = CalcEditKolichPropertiesEditValueChanged
        TabOrder = 2
        Width = 97
      end
      object CalcEditShtuk: TcxCalcEdit
        Left = 175
        Top = 88
        EditValue = '0'
        Enabled = False
        Properties.ImmediatePost = True
        Properties.OnEditValueChanged = CalcEditShtukPropertiesEditValueChanged
        TabOrder = 3
        Width = 97
      end
      object CheckBoxShtuk: TcxCheckBox
        Left = 171
        Top = 49
        Caption = #1096#1090#1091#1082' '#1074' '#1086#1073#1088#1072#1079#1094#1072#1093
        Properties.OnEditValueChanged = CheckBoxShtukPropertiesEditValueChanged
        TabOrder = 4
        Width = 116
      end
    end
    object TabSheet7: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'7'
      ImageIndex = 6
      object cxLabel20: TcxLabel
        Left = 20
        Top = 20
        Caption = #1057#1087#1086#1089#1086#1073' '#1090#1088#1072#1085#1089#1087#1086#1088#1090#1080#1088#1086#1074#1082#1080' '#1075#1088#1091#1079#1072
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object cxLabel21: TcxLabel
        Left = 20
        Top = 55
        Caption = #1058#1088#1072#1085#1089#1087#1086#1088#1090#1085#1086#1077' '#1089#1088#1077#1076#1089#1090#1074#1086
      end
      object ButtonSprTransport: TcxButton
        Left = 509
        Top = 53
        Width = 20
        Height = 19
        Caption = '...'
        TabOrder = 1
        OnClick = ButtonSprTransportClick
        OnExit = ButtonSprTransportExit
      end
      object cxLabel22: TcxLabel
        Left = 20
        Top = 90
        Caption = #1053#1072#1079#1074#1072#1085#1080#1077', '#1085#1086#1084#1077#1088', '#1087#1088#1080#1093#1086#1076'  '
        Properties.WordWrap = True
        Width = 121
      end
      object TextEditNazvTransp: TcxTextEdit
        Left = 203
        Top = 92
        Properties.OnEditValueChanged = TextEditNazvTranspPropertiesEditValueChanged
        TabOrder = 2
        Text = #1085#1077#1090' '#1076#1072#1085#1085#1099#1093
        Width = 326
      end
      object ComboBoxTransport: TcxLookupComboBox
        Left = 203
        Top = 52
        Properties.DropDownRows = 20
        Properties.ImmediatePost = True
        Properties.KeyFieldNames = 'IDTRANSPORT'
        Properties.ListColumns = <
          item
            Caption = #1050#1086#1076' '#1090#1088#1072#1085#1089#1087#1086#1088#1090#1072
            MinWidth = 0
            Width = 0
            FieldName = 'IDTRANSPORT'
          end
          item
            Caption = #1053#1072#1079#1074#1072#1085#1080#1077
            FieldName = 'NAMETRANSPORT'
          end>
        Properties.ListFieldIndex = 1
        Properties.ListOptions.ShowHeader = False
        Properties.ListSource = DataSource11
        EditValue = 0
        Style.BorderStyle = ebsOffice11
        Style.ButtonStyle = btsOffice11
        TabOrder = 0
        Width = 308
      end
    end
    object TabSheet8: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'8'
      ImageIndex = 7
      object cxLabel23: TcxLabel
        Left = 20
        Top = 20
        Caption = #1055#1091#1085#1082#1090' '#1085#1072#1079#1085#1072#1095#1077#1085#1080#1103' '#1075#1088#1091#1079#1072
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object RadioGroupPunkt: TcxRadioGroup
        Left = 10
        Top = 38
        Properties.Items = <
          item
            Caption = #1055#1086#1083#1091#1095#1072#1090#1077#1083#1077#1084' '#1103#1074#1083#1103#1077#1090#1089#1103' '#1091#1082#1072#1079#1072#1085#1085#1099#1081' '#1088#1072#1085#1077#1077' '#1082#1083#1080#1077#1085#1090
          end
          item
            Caption = #1055#1088#1086#1095#1077#1077' '#1084#1077#1089#1090#1086' '#1085#1072#1079#1085#1072#1095#1077#1085#1080#1103' '#1075#1088#1091#1079#1072
          end>
        Properties.OnChange = cxRadioGroup1PropertiesChange
        ItemIndex = 0
        Style.BorderColor = clNone
        Style.BorderStyle = ebsNone
        TabOrder = 0
        Height = 75
        Width = 609
      end
      object TextEditMesto: TcxTextEdit
        Left = 213
        Top = 86
        Enabled = False
        Properties.OnEditValueChanged = TextEditMestoPropertiesEditValueChanged
        TabOrder = 1
        Text = #1085#1077#1090
        Visible = False
        Width = 388
      end
      object CheckBoxExpToRF: TcxCheckBox
        Left = 14
        Top = 159
        Caption = #1056#1086#1089#1089#1080#1081#1089#1082#1072#1103' '#1060#1077#1076#1077#1088#1072#1094#1080#1103
        TabOrder = 3
        Width = 201
      end
      object CheckBoxExpToExp: TcxCheckBox
        Left = 14
        Top = 187
        Caption = #1101#1082#1089#1087#1086#1088#1090
        TabOrder = 4
        Width = 121
      end
      object LabelRegNazn: TcxLabel
        Left = 20
        Top = 129
        Caption = #1055#1088#1086#1092#1080#1083#1100' '#1085#1072#1079#1085#1072#1095#1077#1085#1080#1103
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'MS Sans Serif'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
    end
    object TabSheet9: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'9'
      ImageIndex = 8
      object cxLabel24: TcxLabel
        Left = 20
        Top = 20
        Caption = #1050#1086#1084#1091' '#1074#1099#1076#1072#1077#1090#1089#1103' '#1089#1074#1080#1076#1077#1090#1077#1083#1100#1089#1090#1074#1086' '#1101#1082#1089#1087#1077#1088#1090#1080#1079#1099
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object cxCheckBox1: TcxCheckBox
        Left = 17
        Top = 53
        Caption = 
          #1059#1087#1088#1072#1074#1083#1077#1085#1080#1077' '#1060#1077#1076#1077#1088#1072#1083#1100#1085#1086#1081' '#1089#1083#1091#1078#1073#1099' '#1087#1086' '#1074#1077#1090#1077#1088#1080#1085#1072#1088#1085#1086#1084#1091' '#1080' '#1092#1080#1090#1086#1089#1072#1085#1080#1090#1072#1088#1085#1086#1084#1091 +
          ' '#1085#1072#1076#1079#1086#1088#1091
        State = cbsChecked
        TabOrder = 0
        Width = 640
      end
      object cxCheckBox2: TcxCheckBox
        Left = 17
        Top = 80
        Caption = #1050#1083#1080#1077#1085#1090
        State = cbsChecked
        TabOrder = 1
        Width = 121
      end
      object cxCheckBox3: TcxCheckBox
        Left = 17
        Top = 106
        Caption = #1040#1088#1093#1080#1074' '#1083#1072#1073#1086#1088#1072#1090#1086#1088#1080#1080
        State = cbsChecked
        TabOrder = 2
        Width = 200
      end
    end
    object TabSheet10: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'10'
      ImageIndex = 9
      object cxLabel25: TcxLabel
        Left = 20
        Top = 20
        Caption = #1054#1090' '#1082#1086#1075#1086' '#1087#1086#1089#1090#1091#1087#1080#1083#1080' '#1086#1073#1088#1072#1079#1094#1099
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object cxCheckBox4: TcxCheckBox
        Left = 17
        Top = 53
        Caption = 
          #1059#1087#1088#1072#1074#1083#1077#1085#1080#1077' '#1060#1077#1076#1077#1088#1072#1083#1100#1085#1086#1081' '#1089#1083#1091#1078#1073#1099' '#1087#1086' '#1074#1077#1090#1077#1088#1080#1085#1072#1088#1085#1086#1084#1091' '#1080' '#1092#1080#1090#1086#1089#1072#1085#1080#1090#1072#1088#1085#1086#1084#1091 +
          ' '#1085#1072#1076#1079#1086#1088#1091
        TabOrder = 1
        Width = 640
      end
      object cxCheckBox5: TcxCheckBox
        Left = 17
        Top = 80
        Caption = #1050#1083#1080#1077#1085#1090
        TabOrder = 2
        Width = 121
      end
      object cxCheckBox6: TcxCheckBox
        Left = 17
        Top = 105
        Caption = #1056#1077#1092#1077#1088#1077#1085#1090#1085#1099#1081' '#1094#1077#1085#1090#1088
        TabOrder = 3
        Width = 144
      end
    end
    object TabSheet11: TTabSheet
      Caption = #1044#1080#1072#1083#1086#1075'11'
      ImageIndex = 10
      object LabelInfo: TcxLabel
        Left = 20
        Top = 20
        AutoSize = False
        Caption = #1054#1073#1088#1072#1079#1077#1094' '#1079#1072#1088#1077#1075#1080#1089#1090#1088#1080#1088#1086#1074#1072#1085'!'
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
        Properties.WordWrap = True
        Height = 30
        Width = 277
      end
      object cxLabel26: TcxLabel
        Left = 18
        Top = 61
        AutoSize = False
        Caption = 
          #1056#1077#1075#1080#1089#1090#1088#1072#1094#1080#1103' '#1079#1072#1074#1077#1088#1096#1077#1085#1072'. '#1047#1072#1087#1080#1089#1080' '#1087#1086' '#1087#1088#1086#1090#1086#1082#1086#1083#1091' '#1091#1089#1087#1077#1096#1085#1086' '#1079#1072#1085#1077#1089#1077#1085#1099' '#1074' '#1073#1072 +
          #1079#1091' '#1076#1072#1085#1085#1099#1093'!'
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = []
        Style.IsFontAssigned = True
        Properties.WordWrap = True
        Height = 36
        Width = 551
      end
    end
    object TabSheet12: TTabSheet
      Caption = 'TabSheet12'
      ImageIndex = 11
      object cxLabel2: TcxLabel
        Left = 20
        Top = 20
        Caption = #1057#1074#1077#1076#1077#1085#1080#1103' '#1086#1073' '#1086#1073#1088#1072#1079#1094#1077
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -11
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
      end
      object cxButton2: TcxButton
        Left = 582
        Top = 49
        Width = 20
        Height = 19
        Caption = '...'
        TabOrder = 3
        OnClick = ButtonSprObrazetsClick
        OnExit = ButtonSprObrazetsExit
      end
      object cxLabel7: TcxLabel
        Left = 20
        Top = 77
        AutoSize = False
        Caption = 
          #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077' '#1087#1088#1086#1076#1091#1082#1094#1080#1080' ('#1079#1077#1084#1077#1083#1100#1085#1086#1075#1086' '#1091#1095#1072#1089#1090#1082#1072'), '#1086#1090' '#1082#1086#1090#1086#1088#1086#1075#1086' '#1086#1090#1086#1073#1088#1072#1085 +
          ' '#1086#1073#1088#1072#1079#1077#1094
        Properties.WordWrap = True
        Height = 35
        Width = 213
      end
      object cxLabel8: TcxLabel
        Left = 20
        Top = 120
        Caption = #1050#1083#1072#1089#1089#1080#1092#1080#1082#1072#1094#1080#1103' '#1087#1088#1086#1076#1091#1082#1094#1080#1080
      end
      object ComboBoxKlassOtech: TcxLookupComboBox
        Left = 240
        Top = 119
        Properties.DropDownRows = 20
        Properties.ImmediatePost = True
        Properties.KeyFieldNames = 'IDKLASS'
        Properties.ListColumns = <
          item
            Caption = #1050#1086#1076' '#1082#1083#1072#1089#1089#1080#1092#1080#1082#1072#1094#1080#1080
            MinWidth = 0
            Width = 0
            FieldName = 'IDKLASS'
          end
          item
            Caption = #1053#1072#1079#1074#1072#1085#1080#1077
            FieldName = 'NAMEKLASS'
          end>
        Properties.ListFieldIndex = 1
        Properties.ListOptions.ShowHeader = False
        Properties.ListSource = DataSource7
        EditValue = 0
        Style.BorderStyle = ebsOffice11
        Style.ButtonStyle = btsOffice11
        TabOrder = 5
        Width = 344
      end
      object cxButton4: TcxButton
        Left = 582
        Top = 120
        Width = 19
        Height = 19
        Caption = '...'
        TabOrder = 6
        OnClick = ButtonSprKlassClick
        OnExit = ButtonSprKlassExit
      end
      object cxLabel11: TcxLabel
        Left = 20
        Top = 209
        Caption = #1055#1088#1086#1080#1089#1093#1086#1078#1076#1077#1085#1080#1077
      end
      object cxLabel27: TcxLabel
        Left = 20
        Top = 236
        AutoSize = False
        Caption = #1050#1086#1083#1080#1095#1077#1089#1090#1074#1086' '#1087#1088#1086#1076#1091#1082#1094#1080#1080' ('#1074#1077#1089', '#1096#1090#1091#1082'), '#1087#1083#1086#1097#1072#1076#1100' '#1079#1077#1084#1077#1083#1100#1085#1086#1075#1086' '#1091#1095#1072#1089#1090#1082#1072
        Properties.WordWrap = True
        Height = 34
        Width = 221
      end
      object CalcEditVesOtech: TcxCalcEdit
        Left = 240
        Top = 242
        EditValue = 0.000000000000000000
        Properties.ImmediatePost = True
        Properties.OnChange = cxCalcEdit1PropertiesChange
        TabOrder = 10
        Width = 64
      end
      object ComboBoxVesOtech: TcxLookupComboBox
        Left = 315
        Top = 242
        Properties.DropDownRows = 20
        Properties.ImmediatePost = True
        Properties.KeyFieldNames = 'IDMEASURE'
        Properties.ListColumns = <
          item
            Caption = #1050#1086#1076
            MinWidth = 0
            Width = 0
            FieldName = 'IDMEASURE'
          end
          item
            Caption = #1053#1072#1079#1074#1072#1085#1080#1077
            FieldName = 'NAMEMEASURE'
          end>
        Properties.ListFieldIndex = 1
        Properties.ListOptions.ShowHeader = False
        Properties.ListSource = DataSource9
        EditValue = 0
        Style.BorderStyle = ebsOffice11
        Style.ButtonStyle = btsOffice11
        TabOrder = 11
        Width = 95
      end
      object cxButton6: TcxButton
        Left = 408
        Top = 243
        Width = 20
        Height = 19
        Caption = '...'
        TabOrder = 13
        OnClick = ButtonSprVesClick
        OnExit = ButtonSprVesExit
      end
      object LabelKolvoMest: TcxLabel
        Left = 20
        Top = 277
        Caption = #1050#1086#1083#1080#1095#1077#1089#1090#1074#1086' '#1084#1077#1089#1090
      end
      object CalcEditEdIzmOtech: TcxCalcEdit
        Left = 240
        Top = 276
        EditValue = 0
        Properties.ImmediatePost = True
        Properties.OnEditValueChanged = cxCalcEdit2PropertiesEditValueChanged
        TabOrder = 14
        Width = 64
      end
      object ComboBoxEdIzmOtech: TcxLookupComboBox
        Left = 315
        Top = 276
        Properties.DropDownRows = 20
        Properties.ImmediatePost = True
        Properties.KeyFieldNames = 'IDEDIZM'
        Properties.ListColumns = <
          item
            Caption = #1050#1086#1076
            MinWidth = 0
            Width = 0
            FieldName = 'IDEDIZM'
          end
          item
            Caption = #1053#1072#1079#1074#1072#1085#1080#1077
            FieldName = 'NAMEEDIZM'
          end>
        Properties.ListFieldIndex = 1
        Properties.ListOptions.ShowHeader = False
        Properties.ListSource = DataSource10
        EditValue = 0
        Style.BorderStyle = ebsOffice11
        Style.ButtonStyle = btsOffice11
        TabOrder = 15
        Width = 95
      end
      object cxButton7: TcxButton
        Left = 408
        Top = 277
        Width = 20
        Height = 19
        Caption = '...'
        TabOrder = 16
        OnClick = ButtonSprEdIzmClick
        OnExit = ButtonSprEdIzmExit
      end
      object RadioGroupPartyPole: TcxRadioGroup
        Left = 443
        Top = 203
        Properties.ImmediatePost = True
        Properties.Items = <
          item
            Caption = #1087#1072#1088#1090#1080#1103' '#1087#1088#1086#1076#1091#1082#1094#1080#1080
          end
          item
            Caption = #1087#1086#1083#1077#1074#1086#1077' '#1086#1073#1089#1083#1077#1076#1086#1074#1072#1085#1080#1077
          end>
        Properties.OnEditValueChanged = RadioGroupPartyPolePropertiesEditValueChanged
        ItemIndex = 0
        Style.BorderColor = clNone
        Style.BorderStyle = ebsOffice11
        TabOrder = 18
        Height = 94
        Width = 159
      end
      object cxLabel3: TcxLabel
        Left = 20
        Top = 48
        Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077' '#1086#1073#1088#1072#1079#1094#1072
      end
      object ComboBoxObrazetsOtech: TcxLookupComboBox
        Left = 240
        Top = 48
        Properties.DropDownRows = 20
        Properties.ImmediatePost = True
        Properties.KeyFieldNames = 'IDOBR'
        Properties.ListColumns = <
          item
            Caption = #1050#1086#1076' '#1086#1073#1088#1072#1079#1094#1072
            MinWidth = 0
            Width = 0
            FieldName = 'IDOBR'
          end
          item
            Caption = #1053#1072#1079#1074#1072#1085#1080#1077
            FieldName = 'NAMEOBR'
          end>
        Properties.ListFieldIndex = 1
        Properties.ListOptions.ShowHeader = False
        Properties.ListSource = DataSource5
        EditValue = 0
        Style.BorderStyle = ebsOffice11
        Style.ButtonStyle = btsOffice11
        TabOrder = 1
        Width = 344
      end
      object TextEditProishOtech: TcxTextEdit
        Left = 240
        Top = 208
        Enabled = False
        Properties.OnEditValueChanged = TextEditProishOtechPropertiesEditValueChanged
        StyleDisabled.BorderColor = clWindowFrame
        StyleDisabled.Color = clCaptionText
        StyleDisabled.TextColor = clWindowFrame
        TabOrder = 19
        Text = #1085#1077#1090' '#1076#1072#1085#1085#1099#1093
        Width = 170
      end
      object TextEditKodNasPunkt: TcxTextEdit
        Left = 211
        Top = 208
        Enabled = False
        TabOrder = 20
        Text = #1082#1086#1076' '#1085#1072#1089#1077#1083#1105#1085#1085#1086#1075#1086' '#1087#1091#1085#1082#1090#1072
        Visible = False
        Width = 25
      end
      object cxButton5: TcxButton
        Left = 408
        Top = 209
        Width = 20
        Height = 19
        Caption = '...'
        TabOrder = 9
        OnClick = cxButton5Click
      end
      object MemoSort: TcxMemo
        Left = 19
        Top = 163
        Lines.Strings = (
          #1085#1077#1090' '#1076#1072#1085#1085#1099#1093)
        Properties.OnEditValueChanged = MemoSortPropertiesEditValueChanged
        TabOrder = 21
        Height = 32
        Width = 565
      end
      object cxLabel29: TcxLabel
        Left = 18
        Top = 147
        Caption = 
          #1040#1089#1089#1086#1088#1090#1080#1084#1077#1085#1090' '#1087#1088#1086#1076#1091#1082#1094#1080#1080', '#1082#1086#1083#1080#1095#1077#1089#1090#1074#1086' ('#1085#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077' '#1079#1077#1084#1077#1083#1100#1085#1099#1093' '#1091#1095#1072#1089#1090#1082 +
          #1086#1074', '#1087#1083#1086#1097#1072#1076#1100')'
      end
      object ButtonMemoSize: TcxButton
        Left = 582
        Top = 164
        Width = 19
        Height = 19
        Hint = #1048#1079#1084#1077#1085#1080#1090#1100' '#1088#1072#1079#1084#1077#1088' '#1086#1073#1083#1072#1089#1090#1080' '#1088#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1085#1080#1103
        Caption = '<>'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 23
        OnClick = ButtonMemoSizeClick
      end
      object ComboBoxProdukts: TcxLookupComboBox
        Left = 240
        Top = 84
        Properties.DropDownRows = 20
        Properties.ImmediatePost = True
        Properties.KeyFieldNames = 'IDPARTII'
        Properties.ListColumns = <
          item
            Caption = #1050#1086#1076' '#1087#1072#1088#1090#1080#1080
            MinWidth = 0
            Width = 0
            FieldName = 'IDPARTII'
          end
          item
            Caption = #1053#1072#1079#1074#1072#1085#1080#1077
            FieldName = 'NAMEPARTII'
          end>
        Properties.ListFieldIndex = 1
        Properties.ListOptions.ShowHeader = False
        Properties.ListSource = DataSource6
        EditValue = 0
        Style.BorderStyle = ebsOffice11
        Style.ButtonStyle = btsOffice11
        TabOrder = 24
        Width = 342
      end
      object cxButton3: TcxButton
        Left = 580
        Top = 85
        Width = 20
        Height = 19
        Caption = '...'
        TabOrder = 25
        OnClick = ButtonSprPartiiClick
        OnExit = ButtonSprPartiiExit
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
    RepeatPause = 0
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
    Width = 635
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
  object ButtonExcelTool: TAdvGlowButton
    Left = 22
    Top = 369
    Width = 116
    Height = 33
    Caption = #1060#1086#1088#1084#1072#1090#1080#1088#1086#1074#1072#1085#1080#1077' '#1074' Excel'
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
    TabOrder = 5
    Visible = False
    OnClick = ButtonExcelToolClick
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
  object DataSource1: TDataSource
    DataSet = ADODataSet1
    Left = 33
    Top = 378
  end
  object ADODataSet1: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from _SPETSIALISTS ORDER BY NAMESPETSIALIST;'
    CommandTimeout = 600
    Parameters = <>
    Left = 5
    Top = 378
  end
  object DataSource2: TDataSource
    DataSet = ADODataSet2
    Left = 95
    Top = 378
  end
  object ADODataSet2: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from _KARPUNKTS ORDER BY NAMEKARPUNKT;'
    CommandTimeout = 600
    Parameters = <>
    Left = 67
    Top = 378
  end
  object ADODataSet3: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from _UPRSHNADZ ORDER BY NAMEUPRSHNADZ;'
    CommandTimeout = 600
    Parameters = <>
    Left = 130
    Top = 378
  end
  object DataSource3: TDataSource
    DataSet = ADODataSet3
    Left = 158
    Top = 378
  end
  object ADODataSet4: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from _CLIENTS ORDER BY NAMECLIENT;'
    CommandTimeout = 600
    Parameters = <>
    Left = 192
    Top = 378
  end
  object DataSource4: TDataSource
    DataSet = ADODataSet4
    Left = 220
    Top = 378
  end
  object ADODataSet5: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from _OBRAZTSI ORDER BY NAMEOBR;'
    CommandTimeout = 600
    Parameters = <>
    Left = 78
    Top = 237
  end
  object DataSource5: TDataSource
    DataSet = ADODataSet5
    Left = 106
    Top = 237
  end
  object ADODataSet6: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from _PARTII ORDER BY NAMEPARTII;'
    CommandTimeout = 600
    Parameters = <>
    Left = 143
    Top = 237
  end
  object DataSource6: TDataSource
    DataSet = ADODataSet6
    Left = 171
    Top = 237
  end
  object ADODataSet7: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from _KLASS ORDER BY NAMEKLASS;'
    CommandTimeout = 600
    Parameters = <>
    Left = 207
    Top = 237
  end
  object DataSource7: TDataSource
    DataSet = ADODataSet7
    Left = 235
    Top = 237
  end
  object ADODataSet8: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from _PROISHOZHD ORDER BY NAMEPROISH;'
    CommandTimeout = 600
    Parameters = <>
    Left = 357
    Top = 313
  end
  object DataSource8: TDataSource
    DataSet = ADODataSet8
    Left = 385
    Top = 313
  end
  object ADODataSet9: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from _VES order by NAMEMEASURE;'
    CommandTimeout = 600
    Parameters = <>
    Left = 422
    Top = 313
  end
  object DataSource9: TDataSource
    DataSet = ADODataSet9
    Left = 450
    Top = 313
  end
  object ADODataSet10: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from _EDIZM;'
    CommandTimeout = 600
    Parameters = <>
    Left = 488
    Top = 313
  end
  object DataSource10: TDataSource
    DataSet = ADODataSet10
    Left = 516
    Top = 313
  end
  object ADODataSet11: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from _TRANSPORT order by NAMETRANSPORT;'
    CommandTimeout = 600
    Parameters = <>
    Left = 488
    Top = 341
  end
  object DataSource11: TDataSource
    DataSet = ADODataSet11
    Left = 516
    Top = 341
  end
  object ADOCommand1: TADOCommand
    Connection = FrmMainWindow.ADOConnection1
    Parameters = <>
    Left = 36
    Top = 237
  end
  object ADODataSet12: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandTimeout = 600
    Parameters = <>
    Left = 8
    Top = 237
  end
  object ADODataSet13: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from _DOCVID order by DOCVIDS;'
    CommandTimeout = 600
    Parameters = <>
    Left = 557
    Top = 341
  end
  object DataSource13: TDataSource
    DataSet = ADODataSet13
    Left = 585
    Top = 341
  end
end
