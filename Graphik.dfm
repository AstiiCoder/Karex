object FrmGraph: TFrmGraph
  Left = 288
  Top = 230
  Width = 680
  Height = 438
  BorderStyle = bsSizeToolWin
  Caption = #1043#1088#1072#1092#1080#1082
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object cxGrid1: TcxGrid
    Left = 0
    Top = 66
    Width = 664
    Height = 336
    Align = alClient
    TabOrder = 0
    LookAndFeel.Kind = lfOffice11
    object cxGrid1DBChartView1: TcxGridDBChartView
      DataController.DataSource = DataSource2
      DiagramColumn.Active = True
      ToolBox.CustomizeButton = True
      ToolBox.DiagramSelector = True
      ToolBox.Position = tpBottom
      ToolBox.Visible = tvAlways
    end
    object cxGrid1Level1: TcxGridLevel
      GridView = cxGrid1DBChartView1
    end
  end
  object AdvToolBarPager1: TAdvToolBarPager
    Left = 0
    Top = 0
    Width = 664
    Height = 66
    ActivePage = AdvPage1
    Caption.Visible = False
    Caption.Height = 0
    CaptionButtons = []
    TabGroups = <>
    TabSettings.EndMargin = 0
    TabSettings.Height = 4
    Persistence.Location = plRegistry
    Persistence.Enabled = False
    ToolBarStyler = FrmMainWindow.AdvToolBarFantasyStyler1
    PageLeftMargin = 0
    PageRightMargin = 0
    OptionPicture.Data = {
      424DA20400000000000036040000280000000900000009000000010008000000
      00006C0000000000000000000000000100000000000000000000000080000080
      000000808000800000008000800080800000C0C0C000C0DCC000F0CAA6000020
      400000206000002080000020A0000020C0000020E00000400000004020000040
      400000406000004080000040A0000040C0000040E00000600000006020000060
      400000606000006080000060A0000060C0000060E00000800000008020000080
      400000806000008080000080A0000080C0000080E00000A0000000A0200000A0
      400000A0600000A0800000A0A00000A0C00000A0E00000C0000000C0200000C0
      400000C0600000C0800000C0A00000C0C00000C0E00000E0000000E0200000E0
      400000E0600000E0800000E0A00000E0C00000E0E00040000000400020004000
      400040006000400080004000A0004000C0004000E00040200000402020004020
      400040206000402080004020A0004020C0004020E00040400000404020004040
      400040406000404080004040A0004040C0004040E00040600000406020004060
      400040606000406080004060A0004060C0004060E00040800000408020004080
      400040806000408080004080A0004080C0004080E00040A0000040A0200040A0
      400040A0600040A0800040A0A00040A0C00040A0E00040C0000040C0200040C0
      400040C0600040C0800040C0A00040C0C00040C0E00040E0000040E0200040E0
      400040E0600040E0800040E0A00040E0C00040E0E00080000000800020008000
      400080006000800080008000A0008000C0008000E00080200000802020008020
      400080206000802080008020A0008020C0008020E00080400000804020008040
      400080406000804080008040A0008040C0008040E00080600000806020008060
      400080606000806080008060A0008060C0008060E00080800000808020008080
      400080806000808080008080A0008080C0008080E00080A0000080A0200080A0
      400080A0600080A0800080A0A00080A0C00080A0E00080C0000080C0200080C0
      400080C0600080C0800080C0A00080C0C00080C0E00080E0000080E0200080E0
      400080E0600080E0800080E0A00080E0C00080E0E000C0000000C0002000C000
      4000C0006000C0008000C000A000C000C000C000E000C0200000C0202000C020
      4000C0206000C0208000C020A000C020C000C020E000C0400000C0402000C040
      4000C0406000C0408000C040A000C040C000C040E000C0600000C0602000C060
      4000C0606000C0608000C060A000C060C000C060E000C0800000C0802000C080
      4000C0806000C0808000C080A000C080C000C080E000C0A00000C0A02000C0A0
      4000C0A06000C0A08000C0A0A000C0A0C000C0A0E000C0C00000C0C02000C0C0
      4000C0C06000C0C08000C0C0A000F0FBFF00A4A0A000808080000000FF0000FF
      000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00FDFDFDFDFDFFFFFFFF00
      0000FDFDFDFDE3E3E3E3FF000000FDE3FFFDFDE3E3E3FF000000FDE3FFFDFDE3
      E3E3FF000000FDE3FFFDE3FFFDE3FD000000FDE3FFFDFDFDFDFDFD000000FDE3
      FFFFFFFFFFFDFD000000FDE3E3E3E3E3E3FDFD000000FDFDFDFDFDFDFDFDFD00
      0000}
    OptionDisabledPicture.Data = {
      424DA20400000000000036040000280000000900000009000000010008000000
      00006C0000000000000000000000000100000000000000000000000080000080
      000000808000800000008000800080800000C0C0C000C0DCC000F0CAA6000020
      400000206000002080000020A0000020C0000020E00000400000004020000040
      400000406000004080000040A0000040C0000040E00000600000006020000060
      400000606000006080000060A0000060C0000060E00000800000008020000080
      400000806000008080000080A0000080C0000080E00000A0000000A0200000A0
      400000A0600000A0800000A0A00000A0C00000A0E00000C0000000C0200000C0
      400000C0600000C0800000C0A00000C0C00000C0E00000E0000000E0200000E0
      400000E0600000E0800000E0A00000E0C00000E0E00040000000400020004000
      400040006000400080004000A0004000C0004000E00040200000402020004020
      400040206000402080004020A0004020C0004020E00040400000404020004040
      400040406000404080004040A0004040C0004040E00040600000406020004060
      400040606000406080004060A0004060C0004060E00040800000408020004080
      400040806000408080004080A0004080C0004080E00040A0000040A0200040A0
      400040A0600040A0800040A0A00040A0C00040A0E00040C0000040C0200040C0
      400040C0600040C0800040C0A00040C0C00040C0E00040E0000040E0200040E0
      400040E0600040E0800040E0A00040E0C00040E0E00080000000800020008000
      400080006000800080008000A0008000C0008000E00080200000802020008020
      400080206000802080008020A0008020C0008020E00080400000804020008040
      400080406000804080008040A0008040C0008040E00080600000806020008060
      400080606000806080008060A0008060C0008060E00080800000808020008080
      400080806000808080008080A0008080C0008080E00080A0000080A0200080A0
      400080A0600080A0800080A0A00080A0C00080A0E00080C0000080C0200080C0
      400080C0600080C0800080C0A00080C0C00080C0E00080E0000080E0200080E0
      400080E0600080E0800080E0A00080E0C00080E0E000C0000000C0002000C000
      4000C0006000C0008000C000A000C000C000C000E000C0200000C0202000C020
      4000C0206000C0208000C020A000C020C000C020E000C0400000C0402000C040
      4000C0406000C0408000C040A000C040C000C040E000C0600000C0602000C060
      4000C0606000C0608000C060A000C060C000C060E000C0800000C0802000C080
      4000C0806000C0808000C080A000C080C000C080E000C0A00000C0A02000C0A0
      4000C0A06000C0A08000C0A0A000C0A0C000C0A0E000C0C00000C0C02000C0C0
      4000C0C06000C0C08000C0C0A000F0FBFF00A4A0A000808080000000FF0000FF
      000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00FDFDFDFDFDFFFFFFFF00
      0000FDFDFDFD07070707FF000000FD07FFFDFD070707FF000000FD07FFFDFD07
      0707FF000000FD07FFFD07FFFD07FD000000FD07FFFDFDFDFDFDFD000000FD07
      FFFFFFFFFFFDFD000000FD070707070707FDFD000000FDFDFDFDFDFDFDFDFD00
      0000}
    TabOrder = 1
    object AdvPage1: TAdvPage
      Left = 4
      Top = 4
      Width = 656
      Height = 57
      Caption = 'AdvPage1'
      TabVisible = False
      object AdvToolBar1: TAdvToolBar
        Left = 3
        Top = 3
        Width = 427
        Height = 51
        AllowFloating = True
        AutoPositionControls = False
        AutoSize = False
        Caption = 'Untitled'
        CaptionFont.Charset = DEFAULT_CHARSET
        CaptionFont.Color = clWindowText
        CaptionFont.Height = -11
        CaptionFont.Name = 'MS Sans Serif'
        CaptionFont.Style = []
        CaptionPosition = cpBottom
        CaptionAlignment = taCenter
        CompactImageIndex = -1
        ShowRightHandle = False
        TextAutoOptionMenu = 'Add or Remove Buttons'
        TextOptionMenu = 'Options'
        ToolBarStyler = FrmMainWindow.AdvToolBarFantasyStyler1
        ParentOptionPicture = True
        ParentShowHint = False
        ToolBarIndex = 0
        object AdvToolBarSeparator1: TAdvToolBarSeparator
          Left = 128
          Top = 1
          Width = 10
          Height = 49
          LineColor = clBtnShadow
        end
        object ButtonSave: TAdvGlowButton
          Left = 2
          Top = 2
          Width = 57
          Height = 48
          Hint = #1057#1086#1093#1088#1072#1085#1080#1090#1100' '#1074' '#1092#1072#1081#1083
          Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
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
            424D360C00000000000036000000280000002000000020000000010018000000
            0000000C0000C40E0000C40E0000000000000000000000FFFF00FFFF00FFFF00
            FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
            00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
            FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
            00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
            FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
            00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
            FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FFFFC2C2C2BFBFBFBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBE
            BEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBE
            BEBEBEBEBEBEBEBEBEBEBEBEBECBCBCB00FFFF00FFFF00FFFF00FFFF00FFFF00
            FFFFB99D9E9C5061A157659B5361893B4B9D9A9A9B9898989393938D8D8D8989
            8A8383847E7E8079797C7472796F6F756E6D756E6D903D4E903D4E903D4E903D
            4E903D4E903D4E903D4E903D4BBEBEBE00FFFF00FFFF00FFFF00FFFF00FFFF00
            FFFFAA7676A7626EB16674B06572A25967EDEDEFECECEF90525DB16673A8606D
            DEDCDCD9D6D7D4D0D1CFCCCDCAC6C6C4C0C0D7D4D485404E974A5A974A5AB77A
            84B26774B16773B16673903B4CBEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FA7626EB36976B26875B06874A25967E6E6E9EAE9ED90525DB16673A8606D
            DEDBDED8D6D9D2D0D1CDCBCCC7C3C4C2BDBDD5D1D1833E4D964959974A5AB87A
            84B26875B26874B26774923A4CBEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FBD838BB36B78B36976B26975A25967E2E1E4E6E6E991525EB16874A8606D
            E2E2E4DFDCE0D9D7D9D4D0D2CECBCCC7C4C5D9D6D7813D4C944858964959B87B
            85B36975B36976B36976923B4DBEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FBE848CB36C79B36B78B26976A25967DCDBDEE2E0E3925460B26A75A8606D
            E6E7EAE3E2E5DFDEE1D9D8DAD5D1D4CFCBCDDCDBDB7F3B49924556944858B87B
            86B36976B36976B36976913D4EBEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FBF858CB46C7AB36C7AB36A78A25967D6D5D7DCDADC925460B36A78A8606D
            EBEAEDE7E6EBE3E3E6DFDFE1DAD9DBD6D4D6E2DFE17D3948904354924556BA7C
            86B26975B36976B36976933E51BEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FBF858EB56D7CB46C7BB46C79A25967D0CDCFD7D4D6935561884D59884D59
            E7E7EBEBEBEEE7E7EBE5E4E6E1DFE2DBD9DBE5E4E57C38478D4152904354B97C
            87B26774B26875B26875934051BEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FBE868EB56E7CB66E7BB56C7AA25967CAC6C6CFCCCED5D2D5DBD8DBDFDEE1
            E4E2E5E7E7EBEBEAEEE7E9ECE5E4E7E1DFE3EAE7EA7C38477C38477D3948BB7D
            87B06774B16773B26774934352BEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FC1878EB66F7DB66E7DB56E7BB46C7BA25967A25967A25967A25967A25967
            A25967A25967A25967A25967A25967A25967A25967A25967A25967B06F7AB064
            72B06572B06573B16673954354BEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FC08891B7707EB7707DB56E7CB56D7CB46C7BB46C7AB36B78B26976B26975
            B16875B16673B06673B06472B06271B06270B06270AE6270AD616FAD616FAE62
            70B06471B06471B06572964355BEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FC08891B7707EB7707DB56E7CB56D7CB46C7BB46C7AB36B78B26976B26975
            B16875B16673B06472B06271B06271B0616FAE616FAE6270AD616FAD616FAD61
            6FAE616FAE616FB06271964455BEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FC08992B7737FB56D7A9E53659E53659E53659E53659E53659E53659E5365
            9E53659E53659E53659E53659E53659E53659E53659E53659E53659E53659E53
            659E5365AE606FAE6170964657BEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FC18A93B87480A76470E9DEDEFEFFFEFEFFFEFEFFFEFEFFFEFEFFFEFEFFFE
            FEFEFDFDFEFDFDFDFCFCFCFCFBFCFBFBFBFBF9F9F9F8F8F8F8F8F8F7F7F7E5CC
            CC9E5365B06270AE616F964757BEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FC28B94B97581A76470FEFFFEFCFCFBFBFBF9F9FBF9F9F9F9F8F9F8F8F8F8
            F8F8F8F8F8F7F7F7F7F7F7F6F6F6F6F6F6F6F6F6F6F6F6F6F6F5F5F5F5F5F8F8
            F89E5365B26472AE606F974858BEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FC38B94BA7682A76470FEFFFEDEDEDEDBDBDBD9D9D9D7D8D7D6D6D5D5D5D5
            D2D2D1D0D1D0CFCFCECDCECDCCCCCCCBCBCBCACACACACACACACACACACACAF8F9
            F89E5365B36573AE606E98495BBEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FC38C95BA7683A76470FEFFFEFCFCFCFCFCFBFBFBFBFBFBF9FBF9F9F9F9F9
            F8F8F8F8F8F8F8F8F8F8F8F7F7F7F7F7F7F7F6F6F6F6F6F6F6F6F6F6F6F5FBFB
            FB9E5365B26975AD606E994B5BBEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FC38E96BB7884A76470FEFFFEE1E1E1E0E0E0DFDFDEDCDCDCDAD9DAD8D8D8
            D6D6D6D5D5D5D2D4D4D1D0D0CFCFCFCDCECECCCDCCCBCBCBCACACACACACAFCFC
            FB9E5365B36A76AD606E9B4E5DBEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FC38E96BB7984A76470FEFFFEFCFCFCFCFCFCFCFCFCFCFCFBFBFBFBFBFBF9
            F9F9F9F9F9F8F8F8F8F8F8F8F8F8F7F7F8F7F7F7F7F7F6F6F6F6F6F6F6F6FCFC
            FC9E5365B36C79AE616F9C4E5EBEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FC48E97BB7984A76470FEFFFEE3E3E3E3E2E3E1E1E1E0E0E0DEDFDFDCDCDC
            DADADAD8D8D8D7D7D7D5D5D5D4D4D4D1D1D1CFD0D0CECECECCCCCCCBCBCBFDFD
            FC9E5365B56C79B062709C5261BEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FC58E97BB7984A76470FEFFFEFDFDFDFDFDFCFCFDFCFCFCFCFCFCFBFBFCFB
            FBFBFBFBFBF9F9F9F9F9F9F8F8F8F8F8F8F8F8F8F8F7F7F7F7F7F7F6F6F6FDFE
            FD9E5365B66E7BB064729D5463BEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FC59298BB7984A76470FEFFFEE3E3E3E3E3E3E3E3E3E2E2E3E1E2E2E1E0E0
            DFDFDFDEDCDCDBDBDBD9D9D9D7D7D8D6D6D5D5D4D5D2D1D2D0D0D0CECFCEFEFE
            FE9E5365B76F7CB065729D5563BEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FC59298BB7984A76470FEFFFEFEFEFDFEFEFDFDFDFCFDFDFCFCFDFCFCFCFC
            FCFCFBFCFCFBFBFBF9FBFBF9F9F9F8F8F8F8F8F8F8F8F8F8F8F8F7F7F7F7FEFF
            FE9E5365AB757EB166749E5563BEBEBE00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FC59298BB7984A76470FEFFFEE3E3E3E3E3E3E3E3E3E3E3E3E3E3E3E3E3E3
            E2E2E2E1E1E0DFE0E0DEDEDEDCDBDBD9D9DAD8D7D8D6D6D6D5D5D5D2D4D2FEFF
            FE9E536590636A954E5CA15563BFBFBF00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FC59298BB7984A76470FEFFFEFEFFFEFEFEFEFEFEFDFEFEFDFDFDFCFDFDFC
            FCFDFCFCFCFCFCFCFBFBFBFBFBFBF9F9F9F9F9F9F8F9F9F8F8F8F8F8F8F8FEFF
            FE9E5365A86A74954F5DA25966BFBFBF00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FC48C95BB7984A76470FEFFFEFEFFFEFEFFFEFEFFFEFEFFFEFEFFFEFEFFFE
            FEFFFEFEFFFEFEFFFEFEFFFEFEFFFEFEFFFEFEFFFEFEFFFEFEFFFEFEFFFEFEFF
            FE9E5365C17481C17481A25A67CCCCCC00FFFF00FFFF00FFFF00FFFF00FFFFA6
            5F6FA65F6FA65F6FA65F6FD5D1D1D5D1D1D5D1D1D5D1D1D5D1D1D5D1D1D5D1D1
            D5D1D1D5D1D1D5D1D1D5D1D1D5D1D1D5D1D1D5D1D1D5D1D1D5D1D1D5D1D1D5D1
            D19E53659E53659E53659E536500FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
            00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
            FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
            FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
            00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
            FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF}
          Transparent = True
          TabOrder = 0
          Appearance.BorderColorDisabled = clBtnShadow
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
          Appearance.ColorMirrorDisabled = clBtnFace
          Appearance.ColorMirrorDisabledTo = 15921906
          Enabled = False
          Layout = blGlyphTop
        end
        object ButtonCreateGraph: TAdvGlowButton
          Left = 362
          Top = 2
          Width = 59
          Height = 48
          Hint = #1055#1086#1089#1090#1088#1086#1080#1090#1100' '#1075#1088#1072#1092#1080#1082' '#1087#1086' '#1087#1072#1088#1072#1084#1077#1090#1088#1072#1084' '#1086#1089#1100'/'#1079#1085#1072#1095#1077#1085#1080#1077
          Caption = #1055#1086#1089#1090#1088#1086#1080#1090#1100
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
            424D360C00000000000036000000280000002000000020000000010018000000
            0000000C0000C40E0000C40E000000000000000000000000FF69676766666668
            6666686666686666686666686666686666686666686666666666666666666666
            6666666666666767676565656767676565656666666666666666666666666666
            666666666767676866666866666967676967679F9A9B0000FF33333334363639
            39393D3D3D4242424747474C4C4C5252525858585C5C5C636363696969707070
            7575757B7B7B7F7F7F8686868A8A8A9191919696969C9C9CA2A2A2A7A7A7ADAD
            ADB3B3B3B9B9B9BEBEBEC2C2C2C7C7C7C8C8C86967670000FF33323437383639
            39393B3D3D4141414947474C4C4C5052535957575D5E5C666463676A686D6F70
            7677757A7A7A7E817F8584888A8B89908F919896959A9C9DA0A3A1A5A8A6ABAD
            ADB3B4B2B9B8BABFBDBDC0C2C2C5C7C7C9C8CC6967660000FF3332343836363B
            3938413C3D4141414947474C4C4C5252525756585D5D5D606262696969706D6F
            7575757D7B7B7F807E848484888A8A8E90909496969E9A9FA2A1A3AAA8A8ADAD
            ADB3B3B3B8B8B8BBBEBCC4C0C5C6C7C5CBCCCA6768660000FF64666668676365
            6768626766666666666567686666666666636566636664666765666666686666
            67666A67686665676868646A666666686469656565636A636565656A65676366
            646866656665676865676866656468636866669D9B9B0000FF0000FF947C76C6
            9E92AB6D5BAC6F5BAE6E5CA96E5BAB6D5BA065514B966393DBAD1BAD4E16AE4E
            18AF4F16AD4D18B04D15A0484A94AE94E0FD2AC5FC29C3FE2EC5FE2DC5FF2CC2
            FE2BC5FF2AC4FF2CC4FD2DC3FF4C93B40000FF0000FF0000FF0000FF997D76DC
            C2BBB17361B07260B27462B27462AF7562A2685546966192DEAE19B55015B453
            19B5501AB65117B55118A5484A93AFA5E8FD34CFFF34CEFF31CCFD33CDFF34CE
            FF33CFFD36D1FF33CEFF33CCFF4F92B30000FF0000FF0000FF0000FF987C75E0
            CDC6B57A67B57B68B57A6AB47966B77967A56E5F499565A4E6BD19BA5418BD54
            17BD5219BD5219BC5314AA4C4A94AEBDF1FD3DD9FC3ED9FF3ED9FF3AD9FD3CDA
            FE3BDAFF3CDBFD3AD9FD3DD9FD4F94B50000FF0000FF0000FF0000FF977D76ED
            E1DBBD826EBC816EBC816EBC826FBC816DA774644A9764AFEDC91BC65917C756
            19C4581AC55917C65818B3524D95ADD5F7FD49E3FC49E0FB4DE0FF49E2FD4BE1
            FF4CDFFF4AE0FE4BE1FF4BE1FF4D94B60000FF0000FF0000FF0000FF997C73F9
            F1F2C28777C08875C08877C38777BE8975AB7A6A4A9863CAF3D820CB5E21CC5F
            1FCC5E1DCC5E21CD5D1BB9554B95ADE4FBFD5FE7FF5FE8FE5FE9FF5FE8FE61E8
            FE62E9FF5FE9FF5FE7FF5FE8FE5093B40000FF0000FF0000FF0000FF987C75FC
            F9F5C88D7DC68E7DC68E7BC68E7DC78F7EB480704B9964DCF8E42BD46729D468
            2AD56829D4682AD66626BF5C4E96AEE7FAFF7BEDFE7AEFFE77EDFE7CEEFF79EC
            FF7AEFFE7AEEFF7DEFFF7CEFFC4D94B50000FF0000FF0000FF0000FF987C75FD
            F9F8CD9485CD9584CC9685CB9584CE9685B985784B9964E0FCE937D97337DC73
            38DB7237DC7536DC7130C6684C95B1EFFBFF97F3FE99F3FE95F3FF95F3FF98F1
            FF96F3FC98F4FF95F3FF98F4FF4F90B60000FF0000FF0000FF0000FF987C75FB
            F9F8D19A8BCF9B8ACF9B8ACF9B8AD19A8BBC8B7D489764DEF9E948E38047E27F
            48E38048E38248E2833ECE754C96AEF1FDFFAFF7FEAFF9FFB0F9FDB2F8FFB1F9
            FFB0F9FDAFF7FFAFF9FFB0F8FF4F93B60000FF0000FF0000FF0000FF987C75FB
            F9F8D5A191D3A090D3A18FD3A090D4A090C09183499865E2FBED59EB8C58E98D
            59EB8C57E98B5AEA914DD47E4D96B2F1FDFDC0FBFDC4FBFFC2FBFDC1FBFFBEFB
            FFBFFCFEC0FBFDC0FCFCC3FCFE4E94B20000FF0000FF0000FF0000FF987C75FE
            FAF9D8A797D8A797D8A797D8A797D8A797C297884C9966E2FBED66F09867F199
            67F19967F09A68F0985DDA8A4E97B3F7FFFFF4FEFEF2FEFFF6FEFEF2FEFEF4FF
            FDF6FEFDF4FDFFF5FCFFF1FFFE5092B50000FF0000FF0000FF0000FF987C75FE
            FAF9DCAD9FDCAE9DDDAF9EDDAEA0DCAE9DC59C8D4C9966E2FBED76F6A575F5A3
            74F5A475F6A577F5A369DE955DB6A35096B34F94B54D92B34C93B44F92B14C93
            B54E92B54E94B14E94B24E92B58DB4C30000FF0000FF0000FF0000FF987C75FE
            FAF9E0B3A5DFB4A5DFB4A5DFB4A5DFB4A5C8A1934B9865E5FCEE81F7AC83F9AE
            82F9AE82F8AD84F7AE75E29D74DE9C75DF9A75DE9B4396620000FF0000FF0000
            FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF977B74FF
            FBFAE5BAABE2B9AAE2B9AAE2B8ABE2B9AACBA6984C9966E6FAED8DFAB58EF9B4
            8EFAB28EF8B58FFAB590FAB58EF9B48DFAB500677A00657A00657A0000FF0000
            FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF987D73FD
            FBFBE6BFB0E4C0B0E6BFB1E8BEB1E4C0B0CEAD9D4A9A65E3FCEE9AFBBB9BFBBB
            9BF9BE97F8BF4F7D80007D974C7A7D9AFABA038AAB64D9F30285A40000FF0000
            FF00697E0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF977C72FD
            FBFBE6C5B5E7C4B6E7C4B6E8C5B7E7C5B5D0B1A2499B67E7FBEEA7F9C5A1F9C3
            A4FBC3A7FBC50182A267DEFB088CB30B93BB52CCE964D9F34FC4E3028BB00181
            9F63D0EC00687D0000FF0000FF0000FF0000FF0000FF0000FF0000FF977C72FE
            FCFCE7CABBEBC9BCE7CABBE9CBBAE9CABBD3B5AA4B9A67E9FCEDADF8CAADFCCB
            ADFAC7ADFCCB84C1CE0B93B667DEFB58CEEF65D6F464D9F362D7F052C6E263D0
            EC007F9F0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF987D73FE
            FCFCECCFC0EECEC1EAD0C0ECD0BFECCFC1D5B9AE4E9C67E8FCEFB3FBCDB7FACD
            B2FAD0B3F9D0B4F9CE109DC75CD2F267DEFB199CBD0587A91393B563D0EC52C6
            E2028BB00000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF977C72FE
            FCFCEFD5C5F1D4C6EED5C5F0D6C6F0D5C7D7C0B1489C65E6FDEEB3F9D0B4FCCE
            B4F9CE007A9D007A9D5CDAFB6BDCFD1AA7D0B5FCCA4399630000FF1494B464D4
            ED51C5E10386A50065880000FF0000FF0000FF0000FF0000FF0000FF977B74FC
            FDFBF1DACBF3D9CBF2DBCCEFDACBF4DACCD7C4B54E9B68E6FDEFEDFEF1EAFBF0
            ECFBF3007FA56DDCFF6DDCFF6EE1FF0B9BCBEAFBF0409B640000FF0486A563D8
            F263D8F263D8F20065880000FF0000FF0000FF0000FF0000FF0000FF987B76FF
            FDFCF2DFD2F4DED2F5E0D1EFDED1F5E0D1D9C8BB8BAD8E4D9E67439864449A64
            3F9A631B97BA007FA55FDCFD70E1FF1FADD745996388B6990000FF169BBA63D8
            F451CDE70489AC0065880000FF0000FF0000FF0000FF0000FF0000FF987B76FF
            FDFCF5E3D8F4E2D7F7E4D5F6E5D8F5E5D5DECBC3DBCEC0DECDC0DCCFC193766F
            0000FF0000FF0000FF0194BF62DCFF72E2FF1EA6D50C9FC619A2CB66DCF957CD
            EE0898BD0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF997F73FD
            FDFDF7EADAF7EADCF7E9DDF9E9DCF7E8DFF9E9DCF8EADEF9E9DCF7E9DD99756D
            0000FF0000FF85CCDE059ED272E2FF62D8FF6EDEFF6CDEFE6ADEFF59D5F566DC
            F90B96BB0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF987D73FE
            FEFEF9EEE0FAEEE2FAEEE2FDEDE0FAEEE4FDF0E2FAEEE2FAEFE1FCEEE291766C
            0000FF0000FF0083AB72E2FF059ED20194BF5DD7FC6BDBFF5AD5F80D9CC50D96
            BE66DCF90B93B90000FF0000FF0000FF0000FF0000FF0000FF0000FF997C73FE
            FEFEFCF3E6FAF1E7FAF1E7FCF3E6FBF2E8FBF2E5FBF3E6FBF4E5FCF3E692756C
            0000FF0000FF0000FF0083AB87CDDF0000FF0094C36BDBFF0087B10000FF0000
            FF0F9CC30000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF997C73FF
            FEFDFDF4EAFCF6EBFDF6EDFDF4EAFBF4EBFEF5EBFBF5EAFEF5ECFEF5EB93766D
            0000FF0000FF0000FF0000FF0000FF0000FF0094C30094C30087B10000FF0000
            FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF947D75F9
            FCFFFFFEFDFFFFFFFEFEFEFEFEFEFFFFFFFFFFFFFFFFFFFFFFFFFFFFFB91746F
            0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
            FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FFBCACA698
            7D7392766B92766B92776D92776D92756C93766D92776D95776C91756EB7A69D
            0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
            FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF}
          Transparent = True
          TabOrder = 1
          OnClick = ButtonCreateGraphClick
          Appearance.BorderColorDisabled = clBtnShadow
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
          Appearance.ColorMirrorDisabled = clBtnFace
          Appearance.ColorMirrorDisabledTo = 15921906
          Layout = blGlyphTop
        end
        object ComboBoxOsi: TcxComboBox
          Left = 199
          Top = 4
          TabOrder = 2
          Width = 157
        end
        object ComboBoxVal: TcxComboBox
          Left = 199
          Top = 27
          TabOrder = 3
          Width = 157
        end
        object cxLabel1: TcxLabel
          Left = 141
          Top = 6
          Caption = #1054#1089#1080
        end
        object cxLabel2: TcxLabel
          Left = 141
          Top = 27
          Caption = #1047#1085#1072#1095#1077#1085#1080#1103
        end
        object AdvGlowButton1: TAdvGlowButton
          Left = 62
          Top = 2
          Width = 67
          Height = 48
          Hint = #1050#1086#1087#1080#1088#1086#1074#1072#1090#1100' '#1074' '#1073#1091#1092#1077#1088
          Caption = #1050#1086#1087#1080#1088#1086#1074#1072#1090#1100
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
            424D360C00000000000036000000280000002000000020000000010018000000
            0000000C0000C40E0000C40E000000000000000000000000FF0000FF0000FF00
            00FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF
            0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
            FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF00
            00FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF
            0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
            FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF00
            00FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF
            0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
            FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF00
            00FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FFC2AD9DA59484
            97857384705E76624F76624F76624F76624F76624F76624F76624F76624F7662
            4F76624F76624F76624F76624F76624F0000FF0000FF0000FF0000FF0000FF00
            00FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FFC3AE9EFEFCFC
            FDFCFBFDFBF9FDF9F8FCF8F7FCF8F6FBF7F5FBF7F4FBF5F3F9F5F3FBF4F2F9F3
            F0F9F3F0F8F2EFF9F3EEF8F2EE76624F0000FF0000FF0000FF0000FF0000FF00
            00FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FFC3AE9EFEFCFC
            FEFCFBFDFBF9FDFBF8FDF9F8FCF8F6FCF7F5FCF6F4FBF6F4FBF5F3FBF4F2F9F4
            F2F9F3F0F9F4F0F9F3EFF9F2EF76624F0000FF0000FF0000FF0000FF0000FF00
            00FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FFC3AE9EFDFDFC
            FDFCFBDED7D4DFD7D2DED7D1DFD7D2DFD6D1DED7D1DFD7D0E0D5CFDFD7CEDFD5
            CEDED2CEDFD4CEF9F4F0F9F3F076624F0000FF0000FF0000FF0000FF0000FF00
            00FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FFC3AE9EFEFCFC
            FDFCFBFDFBF9FDFBF9FCF9F8FCF8F7FCF7F6FBF7F5FBF6F4FBF5F4FBF5F3F9F5
            F3FBF5F2F9F4F2FBF4F0F9F4F276624F0000FF0000FF0000FF0000FF0000FF00
            00FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FFC3AE9EFEFDFC
            FEFCFBDED7D4DFD7D2DED7D1DFD7D2DFD6D1DED7D1DFD7D0E0D5CFDFD7CEDFD5
            CEDED2CEDFD4CEFBF4F2FBF4F276624F0000FF0000FF0000FF0000FF0000FF00
            00FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FFC3AE9EFEFCFD
            FEFCFCFDFCFBFDFCF9FEFBF9FDF9F7FCF8F7FCF8F6FCF7F5FCF7F5FBF6F5FCF6
            F4FBF6F3FBF5F3FBF5F2FBF5F276624F0000FF0000FF0000FF0000FF0000FF00
            00FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FFC3AE9EFEFDFD
            FEFDFCDED7D4DFD7D2DED7D1DFD7D2DFD6D1DED7D1DFD7D0E0D5CFDFD7CEDFD5
            CEDED2CEDFD4CEFCF5F4FBF5F476624F0000FF0000FF0000FF0000FF0000FFC2
            AD9DA5948497857384705E76624F76624F76624F76624F76624FC3AE9EFEFDFD
            FEFDFCFDFCFCFEFCFCFDFCFBFDF9F8FDF9F8FCF9F8FCF9F7FDF8F6FCF7F5FCF7
            F5FCF6F5FCF6F4FCF6F4FCF6F476624F0000FF0000FF0000FF0000FF0000FFC3
            AE9EFEFCFCFDFCFBFDFBF9FDF9F8FCF8F7FCF8F6FBF7F5FBF7F4C5B1A2FFFEFE
            FEFEFEDED7D4DFD7D2DED7D1DFD7D2DFD6D1DED7D1DFD7D0E0D5CFDFD7CEDFD5
            CEDED2CEDFD4CEFCF8F7FDF8F676624F0000FF0000FF0000FF0000FF0000FFC3
            AE9EFEFCFCFEFCFBFDFBF9FDFBF8FDF9F8FCF8F6FCF7F5FCF6F4C7B3A3FFFFFE
            FFFEFEFEFEFEFEFDFDFEFDFCFDFCFCFEFCFBFEFCFBFEFBFBFDFBF9FDFBF8FDF9
            F9FDF9F8FCF9F7FDF9F7FCF9F776624F0000FF0000FF0000FF0000FF0000FFC3
            AE9EFDFDFCFDFCFBDED7D4DFD7D2DED7D1DFD7D2DFD6D1DED7D1C9B4A4FFFEFF
            FEFFFEDED9D7DFDAD6DED8D5DFDAD6DFD9D6DED8D5DED7D4FDFBFBFDFBF9FDFB
            F9FDFBF8FCF9F8FDFBF9FDF9F876624F0000FF0000FF0000FF0000FF0000FFC3
            AE9EFEFCFCFDFCFBFDFBF9FDFBF9FCF9F8FCF8F7FCF7F6FBF7F5C9B5A5FEFEFF
            FEFEFFFEFFFEFFFEFEFFFDFDFEFDFDFEFDFCFEFDFCFEFCFCFEFCFBFDFCFBFDFC
            FBFDFBF9FDFCF9FDFBF9FDFBF976624F0000FF0000FF0000FF0000FF0000FFC3
            AE9EFEFDFCFEFCFBDED7D4DFD7D2DED7D1DFD7D2DFD6D1DED7D1CAB6A6FFFEFE
            FFFEFEDEDBDADEDADADEDAD8DEDADADEDAD8DEDAD8DED9D7FEFCFCFEFDFCB9A4
            93A5948493816F7E695776624F76624F0000FF0000FF0000FF0000FF0000FFC3
            AE9EFEFCFDFEFCFCFDFCFBFDFCF9FEFBF9FDF9F7FCF8F7FCF8F6CAB7A7FFFEFE
            FFFEFEFCF6F4FBF6F3FFFEFFFFFEFDFEFEFEFFFDFDFEFDFDFEFDFCFEFCFCBAA5
            95F8EEE9E4D8D0C3B4A876624F0000FF0000FF0000FF0000FF0000FF0000FFC3
            AE9EFEFDFDFEFDFCDED7D4DFD7D2DED7D1DFD7D2DFD6D1DED7D1CBB7A8FFFEFE
            FFFEFEDCDCDCDEDCDBDEDBDADEDCDBDEDBDADEDBDADEDADAFEFDFDFEFDFDBAA6
            96E3D7CFAA9A8A8773610000FF0000FF0000FF0000FF0000FF0000FF0000FFC3
            AE9EFEFDFDFEFDFCFDFCFCFEFCFCFDFCFBFDF9F8FDF9F8FCF9F8CCB8A9FFFEFE
            FFFEFEFFFEFEFFFEFEFFFEFEFFFEFEFEFFFEFFFEFEFEFEFEFEFEFEFEFEFDBDA7
            96B9AA9C9785740000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FFC5
            B1A2FFFEFEFEFEFEDED7D4DFD7D2DED7D1DFD7D2DFD6D1DED7D1CCB9AAFFFEFE
            FFFEFEFFFEFEFFFEFEFFFEFEFFFEFEFFFFFEFFFEFFFFFEFEFFFEFEFEFEFEBEA8
            98A594840000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FFC7
            B3A3FFFFFEFFFEFEFEFEFEFEFDFDFEFDFCFDFCFCFEFCFBFEFCFBCDB9AACCB8A9
            CCB7A8CAB6A6C9B5A5C7B3A3C9B5A5C7B3A3C5B1A2C4B09FC2AD9CBFAA9BBFA9
            990000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FFC9
            B4A4FFFEFFFEFFFEDED9D7DFDAD6DED8D5DFDAD6DFD9D6DED8D5DED7D4FDFBFB
            FDFBF9FDFBF9FDFBF8FCF9F8FDFBF9FDF9F876624F0000FF0000FF0000FF0000
            FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FFC9
            B5A5FEFEFFFEFEFFFEFFFEFFFEFEFFFDFDFEFDFDFEFDFCFEFDFCFEFCFCFEFCFB
            FDFCFBFDFCFBFDFBF9FDFCF9FDFBF9FDFBF976624F0000FF0000FF0000FF0000
            FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FFCA
            B6A6FFFFFFFFFFFFDEDBDADEDADADEDAD8DEDADADEDAD8DEDAD8DED9D7FEFCFC
            FEFDFCB9A493A5948493816F7E695776624F76624F0000FF0000FF0000FF0000
            FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FFCA
            B7A7FFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFEFDFEFEFEFFFDFDFEFDFDFEFDFC
            FEFCFCBAA595F8EEE9E4D8D0C3B4A876624F0000FF0000FF0000FF0000FF0000
            FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FFCB
            B7A8FFFFFFFFFFFFDCDCDCDEDCDBDEDBDADEDCDBDEDBDADEDBDADEDADAFEFDFD
            FEFDFDBAA696E3D7CFAA9A8A8773610000FF0000FF0000FF0000FF0000FF0000
            FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FFCC
            B8A9FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFEFEFEFFFEFFFEFEFEFEFEFEFEFE
            FEFEFDBDA796B9AA9C9785740000FF0000FF0000FF0000FF0000FF0000FF0000
            FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FFCC
            B9AAFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFEFFFFFEFEFFFEFE
            FEFEFEBEA898A594840000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
            FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FFCD
            B9AACCB8A9CCB7A8CAB6A6C9B5A5C7B3A3C9B5A5C7B3A3C5B1A2C4B09FC2AD9C
            BFAA9BBFA9990000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
            FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF00
            00FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF
            0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
            FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF00
            00FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF
            0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF0000
            FF0000FF0000FF0000FF0000FF0000FF0000FF0000FF}
          Transparent = True
          TabOrder = 6
          Appearance.BorderColorDisabled = clBtnShadow
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
          Appearance.ColorMirrorDisabled = clBtnFace
          Appearance.ColorMirrorDisabledTo = 15921906
          Enabled = False
          Layout = blGlyphTop
        end
      end
    end
  end
  object ADODataSet2: TADODataSet
    Connection = FrmMainWindow.ADOConnection1
    CursorType = ctDynamic
    CommandTimeout = 600
    Parameters = <>
    Left = 52
    Top = 356
  end
  object DataSource2: TDataSource
    DataSet = ADODataSet2
    Left = 80
    Top = 356
  end
end
