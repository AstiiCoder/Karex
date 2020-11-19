object FrmNumProtocols: TFrmNumProtocols
  Left = 772
  Top = 209
  BorderStyle = bsToolWindow
  Caption = #1056#1077#1074#1080#1079#1080#1103' '#1085#1086#1084#1077#1088#1086#1074' '#1087#1088#1086#1090#1086#1082#1086#1083#1086#1074
  ClientHeight = 403
  ClientWidth = 219
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
  object cxGrid1: TcxGrid
    Left = 8
    Top = 8
    Width = 201
    Height = 385
    TabOrder = 0
    LookAndFeel.Kind = lfOffice11
    object cxGrid1DBTableView1: TcxGridDBTableView
      NavigatorButtons.ConfirmDelete = False
      DataController.DataSource = DataSource2
      DataController.Summary.DefaultGroupSummaryItems = <>
      DataController.Summary.FooterSummaryItems = <>
      DataController.Summary.SummaryGroups = <>
      OptionsCustomize.ColumnFiltering = False
      OptionsCustomize.ColumnMoving = False
      OptionsCustomize.ColumnSorting = False
      OptionsData.Deleting = False
      OptionsData.Editing = False
      OptionsData.Inserting = False
      OptionsView.GroupByBox = False
    end
    object cxGrid1Level1: TcxGridLevel
      GridView = cxGrid1DBTableView1
    end
  end
  object ADODataSet2: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctDynamic
    CommandTimeout = 600
    Parameters = <>
    Left = 17
    Top = 356
  end
  object DataSource2: TDataSource
    DataSet = ADODataSet2
    Left = 45
    Top = 356
  end
end
