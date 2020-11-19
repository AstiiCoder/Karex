unit Grid1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxStyles, cxCustomData, cxGraphics, cxFilter, cxData,
  cxDataStorage, cxEdit, DB, cxDBData, cxGridLevel, cxClasses, cxControls,
  cxGridCustomView, cxGridCustomTableView, cxGridTableView,  cxGridExportLink,
  ImgList, cxGridCustomPopupMenu, cxGridPopupMenu, AdvToolBar, cxDBLookupComboBox,
  AdvGlowButton, ExtCtrls, AdvPanel, cxGridDBTableView, cxGrid, StdCtrls,
  cxGridStrs, cxGridPopupMenuConsts, ComObj, cxFilterControlStrs, cxEditConsts,
  cxGridBandedTableView, cxGridDBBandedTableView, cxImageComboBox, ADODB;

type
  TFrmGrid1 = class(TForm)
    cxGrid1: TcxGrid;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    cxGridPopupMenu1: TcxGridPopupMenu;
    AdvPanel1: TAdvPanel;
    ButtonOK: TAdvGlowButton;
    ButtonAdd: TDBAdvGlowButton;
    ButtonEdit: TDBAdvGlowButton;
    ButtonDelete: TDBAdvGlowButton;
    AdvToolBarPager1: TAdvToolBarPager;
    AdvPage1: TAdvPage;
    AdvToolBar1: TAdvToolBar;
    ButtonSave: TAdvGlowButton;
    ButtonExportToExcel: TAdvGlowButton;
    SaveDialog1: TSaveDialog;
    ButtonShowColumn: TAdvGlowButton;
    ImageList1: TImageList;
    ButtonFiltr: TAdvGlowButton;
    ToolBarExpertize: TAdvToolBar;
    ButtonProtokol: TAdvGlowButton;
    ButtonSvidetelstvo: TAdvGlowButton;
    ImageList2: TImageList;
    ButtonAkt: TAdvGlowButton;
    ButtonSvvo2: TAdvGlowButton;
    AdvToolBarSeparator1: TAdvToolBarSeparator;
    AdvToolBarSeparator2: TAdvToolBarSeparator;
    cxGrid1DBBandedTableView1: TcxGridDBBandedTableView;
    ButtonSvvo3: TAdvGlowButton;
    ButtonDeleteProtokol: TAdvGlowButton;
    ButtonSwap: TAdvGlowButton;
    ADOCommand1: TADOCommand;
    AdvToolBarSeparator3: TAdvToolBarSeparator;
    AdvPage2: TAdvPage;
    AdvToolBar2: TAdvToolBar;
    ToolBarExcelFonts: TAdvToolBar;
    ButtonSmallFont: TAdvGlowButton;
    ButtonLargeFont: TAdvGlowButton;
    ButtonShowExcel: TAdvGlowButton;
    ButtonDrawBordert: TAdvGlowButton;
    ButtonClearBordert: TAdvGlowButton;
    ButtonGraph: TAdvGlowButton;
    ButtonExpand: TAdvGlowButton;
    ButtonCollaps: TAdvGlowButton;
    ButtonCopy: TAdvGlowButton;
    AdvToolBarSeparator4: TAdvToolBarSeparator;
    AdvToolBar3: TAdvToolBar;
    ButtonOtkazAnaliza: TAdvGlowButton;
    ButtonClear: TAdvGlowButton;
    procedure FormCreate(Sender: TObject);
    procedure ButtonOKClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ButtonExportToExcelClick(Sender: TObject);
    procedure ButtonSaveClick(Sender: TObject);
    procedure ButtonShowColumnClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ButtonFiltrClick(Sender: TObject);
    procedure cxGrid1DBTableView1FilterControlDialogShow(Sender: TObject);
    procedure ButtonOtkazAnalizaClick(Sender: TObject);
    function LookUpTextSpetsField(FieldName, FieldNameSpets: string): string;
    function LookUpText(FieldName: string): string;
    procedure ButtonProtokolClick(Sender: TObject);
    procedure ButtonSvidetelstvoClick(Sender: TObject);
    procedure ButtonLargeFontClick(Sender: TObject);
    procedure ButtonSmallFontClick(Sender: TObject);
    procedure ButtonShowExcelClick(Sender: TObject);
    procedure ButtonClearClick(Sender: TObject);
    procedure ButtonDrawBordertClick(Sender: TObject);
    procedure ButtonClearBordertClick(Sender: TObject);
    procedure cxGrid1DBTableView1DblClick(Sender: TObject);
    procedure ButtonGraphClick(Sender: TObject);
    procedure ButtonAktClick(Sender: TObject);
    procedure ButtonSvvo2Click(Sender: TObject);
    procedure ButtonSvvo3Click(Sender: TObject);
    procedure ButtonExpandClick(Sender: TObject);
    procedure ButtonCollapsClick(Sender: TObject);
    procedure cxGrid1DBTableView1CellDblClick(
      Sender: TcxCustomGridTableView;
      ACellViewInfo: TcxGridTableDataCellViewInfo; AButton: TMouseButton;
      AShift: TShiftState; var AHandled: Boolean);
    procedure ButtonDeleteProtokolClick(Sender: TObject);
    procedure ButtonSwapClick(Sender: TObject);
    procedure ButtonCopyClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmGrid1: TFrmGrid1;

implementation

uses MainWindow, Dialog1, cxFilterConsts, Graphik, ProishImpObr,
  ProishObrOtech;

{$R *.dfm}

type
  TNotifyEvent = procedure(Sender: TObject) of object;

var
  ChangeStatus : boolean;
  OldValChange : integer;

function PMMessageDialog(const Msg: string; Cptn : string; DlgType: TMsgDlgType;
   Buttons: TMsgDlgButtons; Captions: array of string): Integer;
var
   aMsgDlg: TForm;
   i: Integer;
   dlgButton: TButton;
   CaptionIndex: Integer;
begin
   aMsgDlg := CreateMessageDialog(Msg, DlgType, Buttons);
   captionIndex := 0;
   for i := 0 to aMsgDlg.ComponentCount - 1 do
   begin
     if (aMsgDlg.Components[i] is TButton) then
     begin
       dlgButton := TButton(aMsgDlg.Components[i]);
       if CaptionIndex > High(Captions) then Break;
       dlgButton.Caption := Captions[CaptionIndex];
       Inc(CaptionIndex);
     end;
   end;
   aMsgDlg.Caption:=Cptn;
   Result := aMsgDlg.ShowModal;
end;

procedure StrToClipbrd(StrValue: string);
var
  hMem: THandle;
  pMem: PChar;
begin
  hMem := GlobalAlloc(GHND or GMEM_SHARE, Length(StrValue) + 1);
  if hMem <> 0 then
  begin
    pMem := GlobalLock(hMem);
    if pMem <> nil then
    begin
      StrPCopy(pMem, StrValue);
      GlobalUnlock(hMem);
      if OpenClipboard(0) then
      begin
        EmptyClipboard;
        SetClipboardData(CF_TEXT, hMem);
        CloseClipboard;
      end
      else
        GlobalFree(hMem);
    end
    else
      GlobalFree(hMem);
  end;
end;

procedure TFrmGrid1.FormCreate(Sender: TObject);
begin
  //Hide;
  //��������� �������
  cxSetResourceString(@scxGridGroupByBoxCaption , '< ���������� ���� ��������� ������� ��� ����������� >');
  cxSetResourceString(@scxGridFilterCustomizeButtonCaption , '��������...');
  cxSetResourceString(@scxGridCustomizationFormCaption , '��������� ��������');
  cxSetResourceString(@scxGridCustomizationFormColumnsPageCaption , '�������');
  cxSetResourceString(@scxGridCustomizationFormBandsPageCaption , '�����������');
  cxSetResourceString(@scxGridNoDataInfoText , '');
  //��������� ����
  cxSetResourceString(@cxSGridSortColumnAsc , '�� �����������');
  cxSetResourceString(@cxSGridSortColumnDesc , '�� ��������');
  cxSetResourceString(@cxSGridClearSorting , '��� ����������');
  cxSetResourceString(@cxSGridGroupByThisField , '������������ �� ����� ����');
  cxSetResourceString(@cxSGridRemoveThisGroupItem , '������� �� �����������');
  cxSetResourceString(@cxSGridGroupByBox , '������� �����������');
  cxSetResourceString(@cxSGridAlignmentSubMenu , '������������');
  cxSetResourceString(@cxSGridAlignLeft , '�� ������ ����');
  cxSetResourceString(@cxSGridAlignRight , '�� ������� ����');
  cxSetResourceString(@cxSGridAlignCenter , '�� ������');
  cxSetResourceString(@cxSGridRemoveColumn , '������ �������');
  cxSetResourceString(@cxSGridFieldChooser , '��������� ��������');
  cxSetResourceString(@cxSGridBestFit , '����������� ������');
  cxSetResourceString(@cxSGridBestFitAllColumns , '����������� ������ (��� �������)');
  cxSetResourceString(@cxSGridShowFooter , '������� ������');
  cxSetResourceString(@cxSGridShowGroupFooter , '����� �� �������');
  cxSetResourceString(@cxSGridNoneMenuItem , '���');
  //������� ��� ������
  cxSetResourceString(@scxGridCustomizationFormCaption , '���������');
  cxSetResourceString(@scxGridFilterIsEmpty, '<������ ������>');
  //�������� �������, ����� ���� � cxFilterConsts
  cxSetResourceString(@cxSFilterBoxAllCaption,'(���)');
  cxSetResourceString(@cxSFilterBoxCustomCaption,'(�������...)');
  cxSetResourceString(@cxSFilterBoxBlanksCaption,'(������)');
  cxSetResourceString(@cxSFilterBoxNonBlanksCaption, '(�� ������)');
  //������� �������� � ��������
  cxSetResourceString(@cxSFilterOperatorEqual, '�����');
  cxSetResourceString(@cxSFilterOperatorNotEqual , '�� �����');
  cxSetResourceString(@cxSFilterOperatorLess, '������ ���');
  cxSetResourceString(@cxSFilterOperatorLessEqual, '������ ��� �����');
  cxSetResourceString(@cxSFilterOperatorGreater, '������ ���');
  cxSetResourceString(@cxSFilterOperatorGreaterEqual , '������ ��� �����');
  cxSetResourceString(@cxSFilterOperatorLike , '��������');
  cxSetResourceString(@cxSFilterOperatorNotLike , '�� ��������');
  cxSetResourceString(@cxSFilterOperatorBetween , '�����');
  cxSetResourceString(@cxSFilterOperatorNotBetween , '���');
  cxSetResourceString(@cxSFilterOperatorInList , '� ������ ��������');
  cxSetResourceString(@cxSFilterOperatorNotInList , '�� � ������ ��������');
  cxSetResourceString(@cxSFilterAndCaption , '�');
  cxSetResourceString(@cxSFilterOrCaption , '���');
  cxSetResourceString(@cxSFilterNotCaption , '��');
  cxSetResourceString(@cxSFilterBlankCaption , '�����');
  //������
  cxSetResourceString(@cxSFilterOperatorIsNull , '�����');
  cxSetResourceString(@cxSFilterOperatorIsNotNull , '�� �����');
  cxSetResourceString(@cxSFilterOperatorBeginsWith , '���������� �');
  cxSetResourceString(@cxSFilterOperatorDoesNotBeginWith , '�� ���������� �');
  cxSetResourceString(@cxSFilterOperatorEndsWith , '������������� ��');
  cxSetResourceString(@cxSFilterOperatorDoesNotEndWith , '�� ������������� ��');
  cxSetResourceString(@cxSFilterOperatorContains , '�������� � ����');
  cxSetResourceString(@cxSFilterOperatorDoesNotContain , '�� �������� � ����');
  //����������� ��������
  cxSetResourceString(@cxSFilterBoolOperatorAnd, '�');
  cxSetResourceString(@cxSFilterBoolOperatorOr, '���');
  cxSetResourceString(@cxSFilterBoolOperatorNotAnd , '������������� �'); // not all
  cxSetResourceString(@cxSFilterBoolOperatorNotOr , '������������� ���');   // not any
  cxSetResourceString(@cxSFilterRootButtonCaption , '������');
  cxSetResourceString(@cxSFilterAddCondition , '�������� &�������');
  cxSetResourceString(@cxSFilterAddGroup , '�������� &������');
  cxSetResourceString(@cxSFilterRemoveRow , '�&������ ������');
  cxSetResourceString(@cxSFilterClearAll , '������� &��');
  cxSetResourceString(@cxSFilterFooterAddCondition , '�������, ��� ���������� �������');
  cxSetResourceString(@cxSFilterGroupCaption , '��������� ������������������');
  cxSetResourceString(@cxSFilterRootGroupCaption , '<������>');
  cxSetResourceString(@cxSFilterControlNullString , '<�����>');
  cxSetResourceString(@cxSFilterErrorBuilding , '�� ������� ��������� ������');
  //���������� ���� �������
  cxSetResourceString(@cxSFilterDialogCaption , '���������������� ������');
  cxSetResourceString(@cxSFilterDialogInvalidValue , '�� ���������� ��������');
  cxSetResourceString(@cxSFilterDialogUse , '������');
  cxSetResourceString(@cxSFilterDialogSingleCharacter , '���������� ����� ������ ������');
  cxSetResourceString(@cxSFilterDialogCharactersSeries , '���������� ������������������ ������');
  cxSetResourceString(@cxSFilterDialogOperationAnd , '�');
  cxSetResourceString(@cxSFilterDialogOperationOr , '���');
  cxSetResourceString(@cxSFilterDialogRows , '���������� ������ �� ������, �������� �������:');
  //����� ����������� �������
  cxSetResourceString(@cxSFilterControlDialogCaption , '����������� ��������');
  cxSetResourceString(@cxSFilterControlDialogNewFile , 'untitled.flt');
  cxSetResourceString(@cxSFilterControlDialogOpenDialogCaption , '������� ������');
  cxSetResourceString(@cxSFilterControlDialogSaveDialogCaption , '��������� ������ � ����');
  cxSetResourceString(@cxSFilterControlDialogActionSaveCaption , '&��������� ���...');
  cxSetResourceString(@cxSFilterControlDialogActionOpenCaption , '&�������...');
  cxSetResourceString(@cxSFilterControlDialogActionApplyCaption , '&���������');
  cxSetResourceString(@cxSFilterControlDialogActionOkCaption , '��');
  cxSetResourceString(@cxSFilterControlDialogActionCancelCaption , '������');
  //���������
  cxSetResourceString(@cxSDatePopupToday, '�������');
  cxSetResourceString(@cxSDatePopupClear	, '�����');
  //���������
  cxSetResourceString(@scxGridChartCategoriesDisplayText, '������');
  cxSetResourceString(@scxGridChartValueHintFormat, '%s ��� %s - %s');
  cxSetResourceString(@scxGridChartToolBoxDataLevels, '������');
  cxSetResourceString(@scxGridChartToolBoxDataLevelSelectValue, '�������� ��������');
  cxSetResourceString(@scxGridChartToolBoxCustomizeButtonCaption, '�������� ���������');
  cxSetResourceString(@scxGridChartNoneDiagramDisplayText, '�� ���������');
  cxSetResourceString(@scxGridChartColumnDiagramDisplayText, '�����������');
  cxSetResourceString(@scxGridChartBarDiagramDisplayText, '����������');
  cxSetResourceString(@scxGridChartLineDiagramDisplayText, '������');
  cxSetResourceString(@scxGridChartAreaDiagramDisplayText, '�����������');
  cxSetResourceString(@scxGridChartPieDiagramDisplayText, '��������');
  cxSetResourceString(@scxGridChartCustomizationFormSeriesPageCaption, '�����');
  cxSetResourceString(@scxGridChartCustomizationFormSortBySeries, '���������� ��:');
  cxSetResourceString(@scxGridChartCustomizationFormNoSortedSeries, '<��� �����>');
  cxSetResourceString(@scxGridChartCustomizationFormDataGroupsPageCaption, '������ ������');
  cxSetResourceString(@scxGridChartCustomizationFormOptionsPageCaption, '�����');
  cxSetResourceString(@scxGridChartLegend, '�������');
  cxSetResourceString(@scxGridChartLegendKeyBorder, '��������� ���������');
  cxSetResourceString(@scxGridChartPosition, '����������');
  cxSetResourceString(@scxGridChartPositionDefault, '�� ���������');
  cxSetResourceString(@scxGridChartPositionNone, '���');
  cxSetResourceString(@scxGridChartPositionLeft, '�����');
  cxSetResourceString(@scxGridChartPositionTop, '������');
  cxSetResourceString(@scxGridChartPositionRight, '������');
  cxSetResourceString(@scxGridChartPositionBottom, '�����');
  cxSetResourceString(@scxGridChartAlignment, '������������');
  cxSetResourceString(@scxGridChartAlignmentDefault, '�� ���������');
  cxSetResourceString(@scxGridChartAlignmentStart, '������');
  cxSetResourceString(@scxGridChartAlignmentCenter, '�����');
  cxSetResourceString(@scxGridChartAlignmentEnd, '�����');
  cxSetResourceString(@scxGridChartOrientation, '����������');
  cxSetResourceString(@scxGridChartOrientationDefault, '�� ���������');
  cxSetResourceString(@scxGridChartOrientationHorizontal, '�������������');
  cxSetResourceString(@scxGridChartOrientationVertical, '�����������');
  cxSetResourceString(@scxGridChartBorder, '� �����');
  cxSetResourceString(@scxGridChartTitle, '�������');
  cxSetResourceString(@scxGridChartToolBox, '������ ������������');
  cxSetResourceString(@scxGridChartDiagramSelector, '��������� ���������');
  cxSetResourceString(@scxGridChartOther, '������');
  cxSetResourceString(@scxGridChartValueHints, '���������');
  //����������� ����
  cxSetResourceString(@cxSGridAvgMenuItem, '�������');
  cxSetResourceString(@cxSGridCountMenuItem, '���������� �����');
  cxSetResourceString(@cxSGridMaxMenuItem, '��������');
  cxSetResourceString(@cxSGridMinMenuItem, '�������');
  cxSetResourceString(@cxSGridNoneMenuItem	, '��� ������');
  cxSetResourceString(@cxSGridSumMenuItem	, '�����');
  //���. ������ �������
  ToolBarExpertize.Hide;
  ToolBarExcelFonts.Hide;
end;

procedure TFrmGrid1.ButtonOKClick(Sender: TObject);
var
  j:integer;
begin
  j:=0;
  while j<cxGrid1DBTableView1.ColumnCount do
    begin
      if copy(cxGrid1DBTableView1.Columns[j].DataBinding.FieldName,1,2)='ID' then
        begin
          KeyFieldDet:=cxGrid1DBTableView1.Columns[j].DataBinding.FieldName;
          j:=cxGrid1DBTableView1.ColumnCount;
        end;
      inc(j);
    end;
  KeyFieldDet:=FrmMainWindow.ADODataSet1.FieldByName(KeyFieldDet).AsString;
  if KeyFieldDet<>'' then
    TmpObj.EditValue:=StrToInt(KeyFieldDet);
  if TmpObj=FrmDialog1.ComboBoxOrganiz then FrmDialog1.TextEditDogovorNum.SetFocus; 
  close;
end;

procedure TFrmGrid1.FormShow(Sender: TObject);
begin
  ButtonAdd.DataSource:=FrmMainWindow.DataSource1;
  ButtonEdit.DataSource:=FrmMainWindow.DataSource1;
  ButtonDelete.DataSource:=FrmMainWindow.DataSource1;
  if TblRezh=1 then
    begin
      ButtonSwap.Visible:=false;
      ButtonDelete.Enabled:=false;
      cxGrid1DBTableView1.OptionsData.Deleting:=false;
      ButtonOK.Visible:=true;
      ButtonDelete.DataSource:=nil;
    end
  else
    begin
      if UpdStr<>'' then ButtonSwap.Visible:=true;
      if FullDostupBool=true then ButtonSwap.Enabled:=true else ButtonSwap.Enabled:=false;
      ButtonDelete.Enabled:=true;
      cxGrid1DBTableView1.OptionsData.Deleting:=true;
      cxGrid1DBTableView1.OptionsData.Editing:=true;
      cxGrid1DBTableView1.OptionsData.Appending:=true;
      ButtonOK.Visible:=false;
      ButtonAdd.DataSource:=FrmMainWindow.DataSource1;
      ButtonEdit.DataSource:=FrmMainWindow.DataSource1;
      ButtonDelete.DataSource:=FrmMainWindow.DataSource1;
    end;
  //������ ��������� � �������� ���������
  ChangeStatus:=false;
  //if WindowState=wsMinimized then WindowState:=wsMaximized;
end;

procedure TFrmGrid1.ButtonExportToExcelClick(Sender: TObject);
var
  ExcelApp, Workbook : Variant;
  i,j : integer;
  Sheet: Variant;
  bm: string;
begin
  Screen.Cursor:=crHourGlass;
  try
    ExcelApp := GetActiveOleObject('Excel.Application');   //���������, ��� �� ����������� Excel
  except
    on EOLESysError do ExcelApp:= CreateOleObject('Excel.Application');   //���� ���, �� ���������
  end;
  ExcelApp.Application.EnableEvents:= false; // ��� ����� �������
  ExcelApp.Application.DisplayAlerts := false; // ���������� � Excel ��������� �� ���� ����� ���������
  Screen.Cursor:=crDefault;
  // ������� ����� (Workbook)
  Workbook := ExcelApp.WorkBooks.Add;
  Sheet := ExcelApp.Sheets[1];
  //��� �����
  Sheet.Name := '������� �� ���� ������';
  for j:=0 to cxGrid1DBTableView1.VisibleColumnCount-1 do
    Sheet.Cells[1, j+1] := cxGrid1DBTableView1.Columns[j].Caption;
  FrmMainWindow.ADODataSet1.DisableControls;
  bm := FrmMainWindow.ADODataSet1.Bookmark;
  FrmMainWindow.ADODataSet1.First;
  for i:=0 to FrmMainWindow.ADODataSet1.RecordCount-1 do
    begin
      for j:=0 to cxGrid1DBTableView1.ColumnCount-1 do
        Sheet.Cells[i+2, j+1] := FrmMainWindow.ADODataSet1.Fields[j].AsString;
      FrmMainWindow.ADODataSet1.Next;
    end;
  ExcelApp.Visible:=true;
  FrmMainWindow.ADODataSet1.EnableControls;
  FrmMainWindow.ADODataSet1.Bookmark:=bm;
  Screen.Cursor:=crDefault;
end;

procedure TFrmGrid1.ButtonSaveClick(Sender: TObject);
var
  ExcelApp, Workbook : Variant;
begin
  SaveDialog1.Filter := '��������� ���� (*.txt)|*.txt|����� � ������� CSV (*.csv)|*.csv|Web-�������� (*.htm)|*.htm|����� Microsoft Excel (*.xls)|*.xls';
  SaveDialog1.FilterIndex:=4;
  if SaveDialog1.Execute then
    case SaveDialog1.FilterIndex of
        1: begin
                ExportGridToText(SaveDialog1.FileName, cxGrid1, false, True, '', '', '');
           end;
        2: begin
                ExportGridToText(SaveDialog1.FileName, cxGrid1, false, True, '|', '', '');
           end;
        3: begin
                ExportGridToHTML(SaveDialog1.FileName, cxGrid1, false, true, 'html');
           end;
        4: begin
            try
              Screen.Cursor:=crHourGlass;
              ExportGridToExcel(SaveDialog1.FileName, cxGrid1, false, true, true);
              //��������� Excel
              try
                ExcelApp := GetActiveOleObject('Excel.Application');   //���������, ��� �� ����������� Excel
              except
                on EOLESysError do ExcelApp:= CreateOleObject('Excel.Application');   //���� ���, �� ���������
              end;
              ExcelApp.Application.EnableEvents:= false; // ��� ����� �������
              ExcelApp.Application.DisplayAlerts:= false; // ���������� � Excel ��������� �� ���� ����� ���������
              ExcelApp.Visible:=true;
              Screen.Cursor:=crDefault;
              Workbook := ExcelApp.WorkBooks.Open(SaveDialog1.FileName);
              //��� �����
              ExcelApp.WorkBooks[1].WorkSheets[1].Name := '������ �� �������';
              if cxGrid1Level1.GridView=cxGrid1DBBandedTableView1 then
                begin
                    ExcelApp.WorkBooks[1].WorkSheets[1].Cells[1,5].Value:='���������������� ���������';
                    ExcelApp.WorkBooks[1].WorkSheets[1].Range['E1:M1'].Merge;
                    //�/� ������
                    ExcelApp.WorkBooks[1].WorkSheets[1].PageSetup.BlackAndWhite := true;
                    //������� �� ������ ��������
                    ExcelApp.WorkBooks[1].WorkSheets[1].PageSetup.PrintTitleRows := '$1:$3';
                    //����������/���������� � ���������
                    ExcelApp.WorkBooks[1].WorkSheets[1].PageSetup.Zoom := 70;
                    //���������� ����������
                    ExcelApp.WorkBooks[1].WorkSheets[1].PageSetup.Orientation:=2;
                    //�������
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[1].ColumnWidth:= 0.3;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[4].ColumnWidth:= 0.3;
                    ExcelApp.WorkBooks[1].WorkSheets[1].PageSetup.LeftMargin:=30.9;
                    ExcelApp.WorkBooks[1].WorkSheets[1].PageSetup.RightMargin:=10.9;
                    ExcelApp.WorkBooks[1].WorkSheets[1].PageSetup.TopMargin:=10.9;
                    ExcelApp.WorkBooks[1].WorkSheets[1].PageSetup.BottomMargin:=5;
                    ExcelApp.WorkBooks[1].WorkSheets[1].PageSetup.HeaderMargin:=10.9;
                    ExcelApp.WorkBooks[1].WorkSheets[1].PageSetup.FooterMargin:=5;
                    //������ ��������
                    ExcelApp.WorkBooks[1].WorkSheets[1].Rows[1].RowHeight:= 12.75;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Rows[3].RowHeight:= 42.75;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[2].ColumnWidth:= 4.86;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[3].ColumnWidth:= 7.86;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[6].ColumnWidth:= 11.43;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[7].ColumnWidth:= 7.71;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[9].ColumnWidth:= 5.14;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[10].ColumnWidth:= 7.86;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[13].ColumnWidth:= 7.86;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[12].ColumnWidth:= 6.29;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[14].ColumnWidth:= 11.29;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[16].ColumnWidth:= 6.14;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[17].ColumnWidth:= 9.43;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[18].ColumnWidth:= 3.71;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[19].ColumnWidth:= 5.71;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[20].ColumnWidth:= 3.71;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[21].ColumnWidth:= 5.71;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[22].ColumnWidth:= 5.71;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[23].ColumnWidth:= 5.71;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[24].ColumnWidth:= 7.43;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Columns[25].ColumnWidth:= 8.86;
                    ExcelApp.WorkBooks[1].WorkSheets[1].Cells[1,5].HorizontalAlignment:= 3;
                end;
              Screen.Cursor:=crDefault;
            except
              Screen.Cursor:=crDefault;
              showmessage('������ ��� ���������� �����! ���������, �� ������� �� �������� ���� � Excel ����� ����������� �������� ������.');
              exit;
            end;
           end;
  end;
end;

procedure TFrmGrid1.ButtonShowColumnClick(Sender: TObject);
begin
  if cxGrid1.FocusedView is TcxCustomGridTableView then
    with TcxCustomGridTableView(cxGrid1.FocusedView).Controller do
      begin
        Customization := True;
      end;
end;

procedure TFrmGrid1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  TblRezh:=0;
  DefaultVal:='';
  ToolBarExpertize.Visible:=false;
  ToolBarExcelFonts.Visible:=false;
  ButtonGraph.Visible:=true;
  ButtonDeleteProtokol.Visible:=false;
  ButtonDelete.Visible:=true;
  cxGrid1Level1.GridView:=cxGrid1DBTableView1;
  Caption:='������� ���� ������';
end;

procedure TFrmGrid1.ButtonFiltrClick(Sender: TObject);
begin
  cxGrid1DBTableView1.Filtering.RunCustomizeDialog(nil);
end;

procedure TFrmGrid1.cxGrid1DBTableView1FilterControlDialogShow(
  Sender: TObject);
//var
  //frm: TfmFilterControlDialog;
begin
  //frm := Sender as TfmFilterControlDialog;
  //frm.Caption := 'Filter Builder22222';
  //frm.btOpen.Visible := False;
  //frm.btSave.Visible := False;
  //frm.FilterControl.ShowLevelLines := False;
end;

procedure TFrmGrid1.ButtonOtkazAnalizaClick(Sender: TObject);
begin
   try
    cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('OTKAZKL').Index].EditValue:=-1;
   except
    showmessage('���� �� ����� ������� ���������� � ������ ��������������, ��������� �������������� � ������� ������ "�����" ��������!');
   end;
end;

function TFrmGrid1.LookUpTextSpetsField(FieldName, FieldNameSpets: string): string;
begin
  try
    Result:=TcxLookupComboBoxProperties(cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName(FieldName).Index].Properties).ListSource.DataSet.Lookup(TcxLookupComboBoxProperties(cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName(FieldName).Index].Properties).KeyFieldNames, cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName(FieldName).Index].DataBinding.Field.AsString, TcxLookupComboBoxProperties(cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName(FieldName).Index].Properties).ListFieldNames);
    showmessage(TcxLookupComboBoxProperties(cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName(FieldName).Index].Properties).ListFieldNames);
  except
    Result:='';
  end;
end;

function TFrmGrid1.LookUpText(FieldName: string): string;
begin
  try
    Result:=TcxLookupComboBoxProperties(cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName(FieldName).Index].Properties).ListSource.DataSet.Lookup(TcxLookupComboBoxProperties(cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName(FieldName).Index].Properties).KeyFieldNames, cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName(FieldName).Index].DataBinding.Field.AsString, TcxLookupComboBoxProperties(cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName(FieldName).Index].Properties).ListFieldNames);
  except
    Result:='';
  end;
end;

procedure TFrmGrid1.ButtonProtokolClick(Sender: TObject);
var
  s: string;
  i: integer;
  Excel : variant;
begin
      ToolBarExcelFonts.Visible:=true;
      //���������, ������� �� Excel
      try
        ExcelApp := GetActiveOleObject('Excel.Application');
      except
        //���� �������, �� "����������"
        if VarIsEmpty(Excel) = false then
          begin
            Excel.Quit;
            Excel := 0;
          end;
      end;
      //��������� "���" Excel
      try
        ExcelApp := CreateOleObject('Excel.Application');
      except
        ShowMessage('������ ��� ������� MS Excel!');
        exit;
      end;
      ExcelApp.Application.EnableEvents:= false; // ��� ����� �������
      ExcelApp.Application.DisplayAlerts := false; // ���������� � Excel ��������� �� ���� ����� ���������
      ExcelApp.Visible:=true;
      Screen.Cursor:=crDefault;
      Workbook := ExcelApp.WorkBooks.Open(ExtractFilePath(Application.ExeName)+'Templates\��������.xls',0,true);
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].EditValue=0 then
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[3, 8] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].EditValue //����� ���������
      else
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[3, 8] :='�-'+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].EditValue; //����� ��������� �������������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[5, 1] :=LookUpText('DO�TYPE');   ;//   ColumnByFieldName('DO�TYPE').  ;  //�������� ���������
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMETIKETKI').Index].EditValue<>'0' then
        begin
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[12, 2] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMETIKETKI').Index].EditValue;  //����� ���������
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[13, 2] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ETIKETKADATE').Index].EditValue; //���� ������ ��������
        end
      else
        begin
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[12, 1] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[12, 2] :='';  //����� ���������
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[13, 1] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[13, 2] :=''; //���� ������ ��������
        end;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[14, 1] :=LookUpText('UPRSHNADZ'); //���� ��� �/� ����
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[21, 1] :=LookUpText('PUNKTKARRAST'); //����. ������ �� ��������� ��������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[27, 1] :=LookUpText('SPETSIALIST'); //����������
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('SPETSIALIST').Index].EditValue='0' then ExcelApp.WorkBooks[1].WorkSheets[1].Cells[26, 1] := '';
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[31, 2] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('DATEZAYAVKA').Index].EditValue; //���� ������
      //ExcelApp.WorkBooks[1].WorkSheets[1].Cells[32, 1] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].EditValue; //������������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[33, 1] :=LookUpText('ORGANIZATION'); //����.�����������
      //ExcelApp.WorkBooks[1].WorkSheets[1].Cells[36, 1] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].EditValue; //�����
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[5, 3] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('DATEPROTOKOL').Index].EditValue; //���� �������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[7, 4] :=LookUpText('NAMEOBRAZTSA'); //������������ �������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 4] :=LookUpText('NAMEPATRY'); //������������ ������
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('SORT').Index].EditValue<>'���' then
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[24, 4] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('SORT').Index].EditValue //����
      else
        begin
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[23, 4] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[24, 4] :='';
        end;
      try s:=' '+LookUpText('EDVES'); except s:=' '; end;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[28, 4] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VES').Index].EditValue+s;//���
      try s:=' '+LookUpText('EDIZM'); except s:=' '; end;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[30, 4] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KOL').Index].EditValue+s;//����
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[5, 5] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KOLOBR').Index].EditValue; //���-�� ���
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('SHTUKVOBR').Index].EditValue<>'0' then
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[7, 5] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('SHTUKVOBR').Index].EditValue  //���� � ���
      else
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[7, 5] :='';
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[5, 6] :=LookUpText('PROISHOZHD');  //������ ���
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[6, 6] :=LookUpText('VIDTRANSPORT');  //���������
      sleep(300);
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NAZVTRANSPORT').Index].EditValue<>'��� ������' then ExcelApp.WorkBooks[1].WorkSheets[1].Cells[7, 6] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NAZVTRANSPORT').Index].EditValue;  //�����������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[6, 7] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PUNKTNAZNACH').Index].EditValue; //����� ����������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[5, 7] :=' ';
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[11, 7] :=' ';
      FrmMainWindow.ADODataSet2.Active:=false;
      s:=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].EditValue;
      FrmMainWindow.ADODataSet2.CommandText:='SELECT ABBR, ADRESSCLIENT FROM CLIENTS LEFT JOIN ABBREVIATURA ON CLIENTS.ABBREVCLIENT = ABBREVIATURA.IDABBR WHERE (((CLIENTS.IDCLIENT)='+s+'))';
      FrmMainWindow.ADODataSet2.Active:=true;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[32, 1] :=FrmMainWindow.ADODataSet2.FieldByName('ABBR').AsString;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[36, 1] :=FrmMainWindow.ADODataSet2.FieldByName('ADRESSCLIENT').AsString;
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMZAYAVKA').Index].EditValue<>'0' then
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[30, 2] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMZAYAVKA').Index].EditValue //����� ������ �������
      else
        begin
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[30, 1] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[30, 2] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[31, 2] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[29, 1] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[31, 1] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[32, 1] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[33, 1] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[36, 1] :='';
        end;
      //���������� ���������
      FrmMainWindow.ADODataSet2.Active:=false;
      s:='SELECT NAMEOBJLAT, NAMEOBJRUS, SOSTOYANIE, PRIMECH FROM EXP_KAROBJ WHERE (((IDREGISTRATION)='+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue+'))';
      {s:=s+' union all SELECT NAMEOBJLAT, NAMEOBJRUS, SOSTOYANIE, PRIMECH FROM EXP_NEKAROTSRF WHERE (((IDREGISTRATION)='+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue+'))';
      s:=s+' union all SELECT NAMEOBJLAT, NAMEOBJRUS, SOSTOYANIE, PRIMECH FROM EXP_NEKARPROCH WHERE (((IDREGISTRATION)='+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue+'))';
      }
      FrmMainWindow.ADODataSet2.CommandText:=s;
      FrmMainWindow.ADODataSet2.Active:=true;
      FrmMainWindow.ADODataSet2.First;
      s:='';
      for i:=1 to FrmMainWindow.ADODataSet2.RecordCount do
        begin
           s:=s+FrmMainWindow.ADODataSet2.FieldByName('NAMEOBJLAT').AsString+' '+FrmMainWindow.ADODataSet2.FieldByName('NAMEOBJRUS').AsString+' '+FrmMainWindow.ADODataSet2.FieldByName('SOSTOYANIE').AsString+' '+FrmMainWindow.ADODataSet2.FieldByName('PRIMECH').AsString+'  ';
           FrmMainWindow.ADODataSet2.Next;
        end;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[5, 8] :=s;
      //�������� ���� �� ���������
      //ExcelApp.WorkBooks[1].WorkSheets[1].protect('344',true,true,false);
      ButtonLargeFont.Enabled:=true;
      ButtonSmallFont.Enabled:=true;
      ExcelApp.DisplayFullScreen:=true;
      ExcelApp.DisplayFullScreen:=false;
end;

procedure TFrmGrid1.ButtonSvidetelstvoClick(Sender: TObject);
var
  s: string;
  i,k: integer;
begin
      ToolBarExcelFonts.Visible:=true;
      //��������� Excel
      try
        ExcelApp := GetActiveOleObject('Excel.Application');   //���������, ��� �� ����������� Excel
      except
        on EOLESysError do ExcelApp:= CreateOleObject('Excel.Application');   //���� ���, �� ���������
      end;
      ExcelApp.Application.EnableEvents:= false; // ��� ����� �������
      ExcelApp.Application.DisplayAlerts := false; // ���������� � Excel ��������� �� ���� ����� ���������
      ExcelApp.Visible:=true;
      Screen.Cursor:=crDefault;
      Workbook := ExcelApp.WorkBooks.Open(ExtractFilePath(Application.ExeName)+'Templates\�������������.xls',0,true);
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[14, 5] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].EditValue; //����� ���������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[14, 7] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('DATEPROTOKOL').Index].EditValue; //���� ���������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[17, 2] :=LookUpText('UPRSHNADZ'); //���� ��� �/� ����
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 2] :=LookUpText('ORGANIZATION'); //����.�����������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[22, 1] :=LookUpText('NAMEOBRAZTSA'); //������������ �������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[23, 5] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KOLOBR').Index].EditValue; //���-�� ���
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[25, 1] :=LookUpText('NAMEPATRY')+' '+LookUpText('SORT'); //������������ ������, ����
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[26, 1] :=LookUpText('SORT')+' '+ LookUpText('PROISHOZHD')+' '+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VES').Index].EditValue;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[28, 3] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMETIKETKI').Index].EditValue;  //����� ���������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[28, 5] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ETIKETKADATE').Index].EditValue; //���� ������ ��������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[29, 1] :=LookUpText('UPRSHNADZ'); //���� ��� �/� ����
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[30, 1] :=LookUpText('SPETSIALIST'); //����������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[31, 3] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMZAYAVKA').Index].EditValue; //����� ������ �������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[31, 5] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('DATEZAYAVKA').Index].EditValue; //���� ������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[31, 6] :=LookUpText('ORGANIZATION'); //����.�����������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[33, 1] :=LookUpText('NAMEOBRAZTSA');//+', '+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KOL').Index].EditValue; //������������ �������, ���-��
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[35, 1] :=LookUpText('UPRSHNADZ'); //���� ��� �/� ����
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[36, 1] :=LookUpText('ORGANIZATION'); //����.�����������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[38, 1] :=LookUpText('PUNKTNAZNACH');//����� ����������
      FrmMainWindow.ADODataSet2.Active:=false;
      s:=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].EditValue;
      FrmMainWindow.ADODataSet2.CommandText:='SELECT ABBR, ADRESSCLIENT FROM CLIENTS LEFT JOIN ABBREVIATURA ON CLIENTS.ABBREVCLIENT = ABBREVIATURA.IDABBR WHERE (((CLIENTS.IDCLIENT)='+s+'))';
      FrmMainWindow.ADODataSet2.Active:=true;
      s:=FrmMainWindow.ADODataSet2.FieldByName('ABBR').AsString+' '+LookUpText('ORGANIZATION');
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[31, 6] :=s;
      s:=s+FrmMainWindow.ADODataSet2.FieldByName('ADRESSCLIENT').AsString;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[36, 1] :=s;
      //���������� ���������
      FrmMainWindow.ADODataSet2.Active:=false;
      s:='SELECT NAMEOBJLAT, NAMEOBJRUS, SOSTOYANIE, PRIMECH FROM EXP_KAROBJ WHERE (((IDREGISTRATION)='+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue+'))';
      s:=s+' union all SELECT NAMEOBJLAT, NAMEOBJRUS, SOSTOYANIE, PRIMECH FROM EXP_NEKAROTSRF WHERE (((IDREGISTRATION)='+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue+'))';
      s:=s+' union all SELECT NAMEOBJLAT, NAMEOBJRUS, SOSTOYANIE, PRIMECH FROM EXP_NEKARPROCH WHERE (((IDREGISTRATION)='+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue+'))';
      FrmMainWindow.ADODataSet2.CommandText:=s;
      FrmMainWindow.ADODataSet2.Active:=true;
      if FrmMainWindow.ADODataSet2.RecordCount>5 then for i:=0 to FrmMainWindow.ADODataSet2.RecordCount-5 do ExcelApp.WorkBooks[1].WorkSheets[1].Cells[i+42, 1].EntireRow.Insert;
      //������ �����������
      FrmMainWindow.ADODataSet2.Active:=false;
      s:='SELECT NAMEOBJLAT, NAMEOBJRUS, SOSTOYANIE, PRIMECH FROM EXP_KAROBJ WHERE (((IDREGISTRATION)='+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue+'))';
      FrmMainWindow.ADODataSet2.CommandText:=s;
      FrmMainWindow.ADODataSet2.Active:=true;
      FrmMainWindow.ADODataSet2.First;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[40, 1] :='������������ ����������� �������';
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[41, 1] :='    ���';
      k:=FrmMainWindow.ADODataSet2.RecordCount;
      if k<>0 then
        for i:=1 to FrmMainWindow.ADODataSet2.RecordCount do
          begin
           s:=FrmMainWindow.ADODataSet2.FieldByName('NAMEOBJLAT').AsString+' '+FrmMainWindow.ADODataSet2.FieldByName('NAMEOBJRUS').AsString+' '+FrmMainWindow.ADODataSet2.FieldByName('SOSTOYANIE').AsString;
           ExcelApp.WorkBooks[1].WorkSheets[1].Cells[i+40, 1] :=' '+IntToStr(i)+'. '+s;
           FrmMainWindow.ADODataSet2.Next;
          end;
      //������ �������������, ������������� � ��
      if k=0 then k:=k+1;
      FrmMainWindow.ADODataSet2.Active:=false;
      s:='SELECT NAMEOBJLAT, NAMEOBJRUS, SOSTOYANIE, PRIMECH FROM EXP_NEKAROTSRF WHERE (((IDREGISTRATION)='+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue+'))';
      FrmMainWindow.ADODataSet2.CommandText:=s;
      FrmMainWindow.ADODataSet2.Active:=true;
      FrmMainWindow.ADODataSet2.First;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[40+k+1, 1] :='������������ ������������� �������, ������������� � ��';
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[40+k+2, 1] :='    ���';
      k:=k+FrmMainWindow.ADODataSet2.RecordCount;
      if FrmMainWindow.ADODataSet2.RecordCount<>0 then
        for i:=1 to FrmMainWindow.ADODataSet2.RecordCount do
          begin
           s:=FrmMainWindow.ADODataSet2.FieldByName('NAMEOBJLAT').AsString+' '+FrmMainWindow.ADODataSet2.FieldByName('NAMEOBJRUS').AsString+' '+FrmMainWindow.ADODataSet2.FieldByName('SOSTOYANIE').AsString;
           ExcelApp.WorkBooks[1].WorkSheets[1].Cells[i+40+k, 1] :=' '+IntToStr(i)+'. '+s;
           FrmMainWindow.ADODataSet2.Next;
          end;
      //������ �������������, ������
      if k=1 then k:=k+1;
      FrmMainWindow.ADODataSet2.Active:=false;
      s:='SELECT NAMEOBJLAT, NAMEOBJRUS, SOSTOYANIE, PRIMECH FROM EXP_NEKARPROCH WHERE (((IDREGISTRATION)='+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue+'))';
      FrmMainWindow.ADODataSet2.CommandText:=s;
      FrmMainWindow.ADODataSet2.Active:=true;
      FrmMainWindow.ADODataSet2.First;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[40+k+2, 1] :='������������ ������������� �������, ������';
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[40+k+3, 1] :='    ���';
      k:=k+FrmMainWindow.ADODataSet2.RecordCount;
      if FrmMainWindow.ADODataSet2.RecordCount<>0 then
        for i:=1 to FrmMainWindow.ADODataSet2.RecordCount do
          begin
           s:=FrmMainWindow.ADODataSet2.FieldByName('NAMEOBJLAT').AsString+' '+FrmMainWindow.ADODataSet2.FieldByName('NAMEOBJRUS').AsString+' '+FrmMainWindow.ADODataSet2.FieldByName('SOSTOYANIE').AsString;
           ExcelApp.WorkBooks[1].WorkSheets[1].Cells[i+40+k, 1] :=' '+IntToStr(i)+'. '+s;
           FrmMainWindow.ADODataSet2.Next;
          end;
      //�������� ���� �� ���������
      //ExcelApp.WorkBooks[1].WorkSheets[1].protect('344',true,true,false);
      ButtonLargeFont.Enabled:=true;
      ButtonSmallFont.Enabled:=true;
      ExcelApp.DisplayFullScreen:=true;
      ExcelApp.DisplayFullScreen:=false;
      //��������� ��� ���������
      {try
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('EXPCLOSE').Index].EditValue:=-1;
      except
        showmessage('���� �� ����� ������� ���������� � ������ ��������������, ��������� �������������� � ���������� ������ "������������� ������"!');
      end;  }
end;

procedure TFrmGrid1.ButtonLargeFontClick(Sender: TObject);
begin
  try
    ExcelApp.WorkBooks[1].WorkSheets[1].unprotect('344');
    ExcelApp.Selection.Font.Size := ExcelApp.Selection.Font.Size+1;
    ExcelApp.WorkBooks[1].WorkSheets[1].protect('344',true,true,false);
  except
    ButtonLargeFont.Enabled:=false;
  end;
end;

procedure TFrmGrid1.ButtonSmallFontClick(Sender: TObject);
begin
  try
    ExcelApp.WorkBooks[1].WorkSheets[1].unprotect('344');
    ExcelApp.Selection.Font.Size := ExcelApp.Selection.Font.Size-1;
    ExcelApp.WorkBooks[1].WorkSheets[1].protect('344',true,true,false);
  except
    ButtonSmallFont.Enabled:=false;
  end;
end;

procedure TFrmGrid1.ButtonShowExcelClick(Sender: TObject);
begin
  ExcelApp.DisplayFullScreen:=true;
  ExcelApp.DisplayFullScreen:=false;
end;

procedure TFrmGrid1.ButtonClearClick(Sender: TObject);
begin
   try
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('UPRSHNADZ').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PUNKTKARRAST').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('SPETSIALIST').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMETIKETKI').Index].EditValue:='0';
     //cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ETIKETKADATE').Index].EditValue:='null';
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMDOGOVOR').Index].EditValue:='0';
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMZAYAVKA').Index].EditValue:='0';
     //cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('DATEZAYAVKA').Index].EditValue:='';
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NAMEOBRAZTSA').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KLASSIFIKATION').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NAMEPATRY').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('SORT').Index].EditValue:='���';
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PROISHOZHD').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VES').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('EDVES').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KOL').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('EDIZM').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KOLOBR').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('SHTUKVOBR').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VIDTRANSPORT').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NAZVTRANSPORT').Index].EditValue:='��� ������';
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PUNKTNAZNACH').Index].EditValue:='���';
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VUPRSHNADZ').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VCLIENT').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VARCHIV').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PUPRSHNADZ').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PCLIENT').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PREFCENTR').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('EXPCLOSE').Index].EditValue:=0;
     cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('OTKAZKL').Index].EditValue:=0;
   except
    showmessage('���� �� ����� ������� ���������� � ������ ��������������, ��������� �������������� � ��������� ��������!');
   end;
end;

procedure TFrmGrid1.ButtonDrawBordertClick(Sender: TObject);
begin
  try
    ExcelApp.WorkBooks[1].WorkSheets[1].unprotect('344');
    ExcelApp.Selection.Borders[4].LineStyle := 1; //Borders[1] - ����� �������
    ExcelApp.WorkBooks[1].WorkSheets[1].protect('344',true,true,false);
  except
    ButtonLargeFont.Enabled:=false;
  end;
end;

procedure TFrmGrid1.ButtonClearBordertClick(Sender: TObject);
begin
  try
    ExcelApp.WorkBooks[1].WorkSheets[1].unprotect('344');
    ExcelApp.Selection.Borders[4].LineStyle := 0;
    ExcelApp.WorkBooks[1].WorkSheets[1].protect('344',true,true,false);
  except
    ButtonLargeFont.Enabled:=false;
  end;
end;

procedure TFrmGrid1.cxGrid1DBTableView1DblClick(Sender: TObject);
begin
  if TblRezh=1 then ButtonOKClick(Sender);
end;

procedure TFrmGrid1.ButtonGraphClick(Sender: TObject);
var
  i: integer;
begin
  if FrmGraph= NIL then FrmGraph := TFrmGraph.Create(owner);
  FrmGraph.ComboBoxOsi.Properties.Items.Clear;
  FrmGraph.ComboBoxVal.Properties.Items.Clear;
  FrmGraph.ComboBoxOsi.Text:='';
  FrmGraph.ComboBoxVal.Text:='';
  for i:=0 to cxGrid1DBTableView1.ColumnCount-1 do
    begin
     FrmGraph.ComboBoxOsi.Properties.Items.Add(cxGrid1DBTableView1.Columns[i].Caption);
     FrmGraph.ComboBoxVal.Properties.Items.Add(cxGrid1DBTableView1.Columns[i].Caption);
    end;
  FrmGraph.ShowModal;
end;

procedure TFrmGrid1.ButtonAktClick(Sender: TObject);
var
   s: string;
   Excel : variant;
begin
      ToolBarExcelFonts.Visible:=true;
      //���������, ������� �� Excel
      try
        ExcelApp := GetActiveOleObject('Excel.Application');
      except
        //���� �������, �� "����������"
        if VarIsEmpty(Excel) = false then
          begin
            Excel.Quit;
            Excel := 0;
          end;
      end;
      //��������� "���" Excel
      try
        ExcelApp := CreateOleObject('Excel.Application');
      except
        ShowMessage('������ ��� ������� MS Excel!');
        exit;
      end;
      ExcelApp.Application.EnableEvents:= false; // ��� ����� �������
      ExcelApp.Application.DisplayAlerts := false; // ���������� � Excel ��������� �� ���� ����� ���������
      ExcelApp.Visible:=true;
      Screen.Cursor:=crDefault;
      Workbook := ExcelApp.WorkBooks.Open(ExtractFilePath(Application.ExeName)+'Templates\���.xls',0,true);
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].EditValue=0 then
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[6, 4] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].EditValue //����� ���������
      else
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[6, 4] :='�-'+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].EditValue; //����� ��������� �������������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[6, 6] :=Date(); //���� �������
      s :=LookUpText('NAMEOBRAZTSA'); //������������ �������
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('SORT').Index].EditValue<>'���' then
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[8, 1] :=s+', ����: '+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('SORT').Index].EditValue //����, ���� ����
      else
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[8, 1] :=s;
      try s:=' '+LookUpText('EDVES'); except s:=' '; end;
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VES').Index].EditValue<>'0' then
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[9, 1] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VES').Index].EditValue+s //���
      else
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[9, 1] :='';
      try s:=' '+LookUpText('EDIZM'); except s:=' '; end;
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KOL').Index].EditValue<>'0' then
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[9, 2] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KOL').Index].EditValue+s //����
      else
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[9, 2] :='';
      //ExcelApp.WorkBooks[1].WorkSheets[1].Cells[5, 5] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KOLOBR').Index].EditValue; //���-�� ���
      //ExcelApp.WorkBooks[1].WorkSheets[1].Cells[7, 5] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('SHTUKVOBR').Index].EditValue;  //���� � ���

      s:=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].EditValue;
      FrmMainWindow.ADODataSet2.Active:=false;
      FrmMainWindow.ADODataSet2.CommandText:='SELECT ABBR, ADRESSCLIENT FROM CLIENTS LEFT JOIN ABBREVIATURA ON CLIENTS.ABBREVCLIENT = ABBREVIATURA.IDABBR WHERE (((CLIENTS.IDCLIENT)='+s+'))';
      FrmMainWindow.ADODataSet2.Active:=true;
      s:=FrmMainWindow.ADODataSet2.FieldByName('ABBR').AsString+' '+LookUpText('ORGANIZATION'); //����.�����������
      s:=s+', '+FrmMainWindow.ADODataSet2.FieldByName('ADRESSCLIENT').AsString;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[10, 3] :=s; //�����������
      //�������������
      s:='';
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VIDTRANSPORT').Index].EditValue<>'0' then s:=' '+LookUpText('VIDTRANSPORT');
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NAZVTRANSPORT').Index].EditValue<>'��� ������' then s:=s+': '+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NAZVTRANSPORT').Index].EditValue;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[12, 3] :=LookUpText('PROISHOZHD')+' '+s;  //������ ���
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[14, 4] :=LookUpText('PUNKTKARRAST'); //����. ������ �� ��������� ��������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[16, 1] :=LookUpText('NAMEOBRAZTSA'); //������������ �������

      {ExcelApp.WorkBooks[1].WorkSheets[1].Cells[5, 1] :=LookUpText('DO�TYPE');   ;//   ColumnByFieldName('DO�TYPE').  ;  //�������� ���������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[12, 2] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMETIKETKI').Index].EditValue;  //����� ���������

      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[14, 1] :=LookUpText('UPRSHNADZ'); //���� ��� �/� ����

      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[27, 1] :=LookUpText('SPETSIALIST'); //����������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[30, 2] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMZAYAVKA').Index].EditValue; //����� ������ �������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[31, 2] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('DATEZAYAVKA').Index].EditValue; //���� ������
      //ExcelApp.WorkBooks[1].WorkSheets[1].Cells[32, 1] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].EditValue; //������������

      //ExcelApp.WorkBooks[1].WorkSheets[1].Cells[36, 1] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].EditValue; //�����
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[5, 3] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('DATEPROTOKOL').Index].EditValue; //���� �������

      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 4] :=LookUpText('NAMEPATRY'); //������������ ������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[24, 4] :=LookUpText('SORT'); //����
      try s:=' '+LookUpText('EDVES'); except s:=' '; end;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[28, 4] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VES').Index].EditValue+s;//���
      try s:=' '+LookUpText('EDIZM'); except s:=' '; end;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[30, 4] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KOL').Index].EditValue+s;//����
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[5, 5] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KOLOBR').Index].EditValue; //���-�� ���
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[7, 5] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('SHTUKVOBR').Index].EditValue;  //���� � ���

      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[6, 6] :=LookUpText('VIDTRANSPORT');  //���������
      sleep(300);
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[7, 6] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NAZVTRANSPORT').Index].EditValue;  //�����������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[6, 7] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PUNKTNAZNACH').Index].EditValue; //����� ����������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[5, 7] :=' ';
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[11, 7] :=' ';
      FrmMainWindow.ADODataSet2.Active:=false;
      s:=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].EditValue;
      FrmMainWindow.ADODataSet2.CommandText:='SELECT ABBR, ADRESSCLIENT FROM CLIENTS LEFT JOIN ABBREVIATURA ON CLIENTS.ABBREVCLIENT = ABBREVIATURA.IDABBR WHERE (((CLIENTS.IDCLIENT)='+s+'))';
      FrmMainWindow.ADODataSet2.Active:=true;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[32, 1] :=FrmMainWindow.ADODataSet2.FieldByName('ABBR').AsString;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[36, 1] :=FrmMainWindow.ADODataSet2.FieldByName('ADRESSCLIENT').AsString;   }
      //���������� ���������
      {FrmMainWindow.ADODataSet2.Active:=false;
      s:='SELECT NAMEOBJLAT, NAMEOBJRUS, SOSTOYANIE, PRIMECH FROM EXP_KAROBJ WHERE (((IDREGISTRATION)='+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue+'))';
      s:=s+' union all SELECT NAMEOBJLAT, NAMEOBJRUS, SOSTOYANIE, PRIMECH FROM EXP_NEKAROTSRF WHERE (((IDREGISTRATION)='+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue+'))';
      s:=s+' union all SELECT NAMEOBJLAT, NAMEOBJRUS, SOSTOYANIE, PRIMECH FROM EXP_NEKARPROCH WHERE (((IDREGISTRATION)='+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue+'))';
      FrmMainWindow.ADODataSet2.CommandText:=s;
      FrmMainWindow.ADODataSet2.Active:=true;
      FrmMainWindow.ADODataSet2.First;
      for i:=1 to FrmMainWindow.ADODataSet2.RecordCount do
        begin
           s:=FrmMainWindow.ADODataSet2.FieldByName('NAMEOBJLAT').AsString+' '+FrmMainWindow.ADODataSet2.FieldByName('NAMEOBJRUS').AsString+' '+FrmMainWindow.ADODataSet2.FieldByName('SOSTOYANIE').AsString+' '+FrmMainWindow.ADODataSet2.FieldByName('PRIMECH').AsString;
           ExcelApp.WorkBooks[1].WorkSheets[1].Cells[i*2+3, 8] :=s;
           FrmMainWindow.ADODataSet2.Next;
        end;  }
      //�������� ���� �� ���������
      //ExcelApp.WorkBooks[1].WorkSheets[1].protect('344',true,true,false);
      ButtonLargeFont.Enabled:=true;
      ButtonSmallFont.Enabled:=true;
      ExcelApp.DisplayFullScreen:=true;
      ExcelApp.DisplayFullScreen:=false;
end;

procedure TFrmGrid1.ButtonSvvo2Click(Sender: TObject);
var
  s, kl: string;
  Excel : variant;
begin
      ToolBarExcelFonts.Visible:=true;
      //���������, ������� �� Excel
      try
        ExcelApp := GetActiveOleObject('Excel.Application');
      except
        //���� �������, �� "����������"
        if VarIsEmpty(Excel) = false then
          begin
            Excel.Quit;
            Excel := 0;
          end;
      end;
      //��������� "���" Excel
      try
        ExcelApp := CreateOleObject('Excel.Application');
      except
        ShowMessage('������ ��� ������� MS Excel!');
        exit;
      end;
      ExcelApp.Application.EnableEvents:= false; // ��� ����� �������
      ExcelApp.Application.DisplayAlerts := false; // ���������� � Excel ��������� �� ���� ����� ���������
      ExcelApp.Visible:=true;
      Screen.Cursor:=crDefault;
      Workbook := ExcelApp.WorkBooks.Open(ExtractFilePath(Application.ExeName)+'Templates\�������������2.xls',0,true);
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].EditValue='0' then
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[6, 5] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].EditValue //����� ���������
      else
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[6, 5] :='�-'+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].EditValue; //����� ���������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[6, 7] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('DATEPROTOKOL').Index].EditValue; //���� ���������
      //���� ������ �������������
      s:='';
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VUPRSHNADZ').Index].EditValue=true then
        s := '���������� ����������� ������ �� ������������� � ��������������� ������� �� �����-���������� � ������������� �������';
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VCLIENT').Index].EditValue=true then
        if s='' then s:='�������' else s:=s+', �������';
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VARCHIV').Index].EditValue=true then
        if s='' then s:='� ����� ���' else s:=s+', � ����� ���';
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[8, 1]:=s;
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('DO�TYPE').Index].EditValue=0 then
        s:='�����������'
      else
        begin
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[10, 1] :=LookUpText('DO�TYPE');  //������������ ���������
          s :='�'+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMETIKETKI').Index].EditValue;  //����� ���������
          s:=s+' �� '+DateToStr(cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ETIKETKADATE').Index].EditValue); //���� ������ ��������
          s:=s+', ����� ����������: '+LookUpText('SPETSIALIST'); //����������
          s:=s+', '+LookUpText('UPRSHNADZ'); //����. ���������� �/� ����
        end;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[11, 1]:=s;
      //�����������
      kl:=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].EditValue;
      FrmMainWindow.ADODataSet2.Active:=false;
      FrmMainWindow.ADODataSet2.CommandText:='SELECT ABBR, ADRESSCLIENT FROM CLIENTS LEFT JOIN ABBREVIATURA ON CLIENTS.ABBREVCLIENT = ABBREVIATURA.IDABBR WHERE (((CLIENTS.IDCLIENT)='+IntToStr(cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].EditValue)+'))';
      FrmMainWindow.ADODataSet2.Active:=true;
      kl:=FrmMainWindow.ADODataSet2.FieldByName('ABBR').AsString+' '+LookUpText('ORGANIZATION'); //����.�����������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[13, 1]:=s+', '+kl; // ����������� ��� ������
      s:=kl+', '+FrmMainWindow.ADODataSet2.FieldByName('ADRESSCLIENT').AsString;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[22, 1] :=s; //����������� � �������
      //������
      s:='';
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMZAYAVKA').Index].EditValue<>'0' then
        begin
          s:='������ �'+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMZAYAVKA').Index].EditValue; //����� ������ �������
          s:=s+' �� '+DateToStr(cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('DATEZAYAVKA').Index].EditValue); //���� ������
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[13, 1]:=s;
        end
      else
        begin
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[13, 1]:='';
          //ExcelApp.WorkBooks[1].WorkSheets[1].Cells[22, 1] :=''; //����������� � �������
        end;
      //�������������
      s:='';
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VIDTRANSPORT').Index].EditValue<>'0' then s:=' '+LookUpText('VIDTRANSPORT');
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NAZVTRANSPORT').Index].EditValue<>'��� ������' then s:=s+': '+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NAZVTRANSPORT').Index].EditValue;
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].EditValue='0' then
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[20, 6] :=LookUpText('PROISHIMP')+' '+s  //������ ���
      else
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[20, 6] := cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PROISHOTECH').Index].EditValue;
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KOLOBR').Index].EditValue<>'0' then
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 5] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KOLOBR').Index].EditValue //���-�� ���
      else
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 5] :=''; //���-�� ���
      s:='';
      s :=LookUpText('NAMEOBRAZTSA'); //������������ �������
      s:=s+' '+LookUpText('NAMEPATRY'); //������������ ������
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('SORT').Index].EditValue<>'���' then s :=s+', ����: '+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('SORT').Index].EditValue; //����, ���� ����
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[17, 1] :=s;

      try s:=' '+LookUpText('EDVES'); except s:=' '; end;
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VES').Index].EditValue<>'0' then
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[18, 1] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VES').Index].EditValue+s //���
      else
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[18, 1] :='';
      try s:=' '+LookUpText('EDIZM'); except s:=' '; end;
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KOL').Index].EditValue<>'0' then
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[18, 4] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KOL').Index].EditValue+s //����
      else
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[18, 4] :='';
      s:='';
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PUPRSHNADZ').Index].EditValue=true then s:='���������� ����������� ������ �� ������������� � ��������������� ������� �� �����-���������� � ������������� �������';
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PREFCENTR').Index].EditValue=true then
        if s<>'' then
          s:=s+', ��� "������������� ����������� ����� ����������� ������ �� ������������� � ��������������� �������"'
        else
          s:=s+'��� "������������� ����������� ����� ����������� ������ �� ������������� � ��������������� �������"';
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PCLIENT').Index].EditValue=true then
        if s<>'' then
          s:=s+', '+kl
        else
          s:=kl;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[15, 1] :=s;
      //���� ����������� ��������
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('DO�TYPE').Index].EditValue='0' then
        begin
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[10, 1] :='�����������';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[11, 1] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[11, 2] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[11, 3] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[11, 4] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[11, 6] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[12, 1] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[12, 2] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[12, 3] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[13, 1] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[13, 3] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[13, 4] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[13, 5] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[13, 7] :='';
        end;
      //����� ����������
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PUNKTNAZNACH').Index].EditValue<>'���' then
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[21, 4] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PUNKTNAZNACH').Index].EditValue
      else
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[21, 4] :='';
      //��� ������������� - ��������� ������ ����������: �����������, ������������� ...
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].EditValue='1' then
        begin
          FrmMainWindow.ADODataSet2.Active:=false;
          FrmMainWindow.ADODataSet2.CommandText:='SELECT * FROM OTECHMEMOSORT WHERE (((OTECHMEMOSORT.IDREGISTRATION)='+IntToStr(StrToInt(cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue)+1)+'))';
          FrmMainWindow.ADODataSet2.Active:=true;
          if FrmMainWindow.ADODataSet2.RecordCount>0 then
            begin
              ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 1].EntireRow.Insert;
              ExcelApp.WorkBooks[1].WorkSheets[1].Range[ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 1], ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 10]].Merge;
              ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 1].EntireRow.Insert;
              ExcelApp.WorkBooks[1].WorkSheets[1].Range[ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 1], ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 10]].Merge;
            end;
          if copy(FrmMainWindow.ADODataSet2.FieldByName('SORTTMPROCH').AsString,1,10)<>'��� ������' then ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 1] :=FrmMainWindow.ADODataSet2.FieldByName('SORTTMPROCH').AsString;
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[20, 1] :=FrmMainWindow.ADODataSet2.FieldByName('PROISH').AsString;
          //ExcelApp.WorkBooks[1].WorkSheets[1].Cells[i+25, 1]
        end;
      //�������� ���� �� ���������
      //ExcelApp.WorkBooks[1].WorkSheets[1].protect('344',true,true,false);
      ButtonLargeFont.Enabled:=true;
      ButtonSmallFont.Enabled:=true;
      ExcelApp.DisplayFullScreen:=true;
      ExcelApp.DisplayFullScreen:=false;
      {try
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('EXPCLOSE').Index].EditValue:=-1;
      except
        showmessage('���� �� ����� ������� ���������� � ������ ��������������, ��������� �������������� � ���������� ������ "������������� ������"!');
      end; }
end;


procedure TFrmGrid1.ButtonSvvo3Click(Sender: TObject);
var
  s, kl, c: string;
  i,k,m: integer;
  Excel : variant;
begin
      ToolBarExcelFonts.Visible:=true;
      //���������, ������� �� Excel
      try
        ExcelApp := GetActiveOleObject('Excel.Application');
      except
        //���� �������, �� "����������"
        if VarIsEmpty(Excel) = false then
          begin
            Excel.Quit;
            Excel := 0;
          end;
      end;
      //��������� "���" Excel
      try
        ExcelApp := CreateOleObject('Excel.Application');
      except
        ShowMessage('������ ��� ������� MS Excel!');
        exit;
      end;
      ExcelApp.Application.EnableEvents:= false; // ��� ����� �������
      ExcelApp.Application.DisplayAlerts := false; // ���������� � Excel ��������� �� ���� ����� ���������
      ExcelApp.Visible:=true;
      Screen.Cursor:=crDefault;
      Workbook := ExcelApp.WorkBooks.Open(ExtractFilePath(Application.ExeName)+'Templates\�������������3.xls',0,true);
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].EditValue='0' then
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[6, 5] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].EditValue //����� ���������
      else
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[6, 5] :='�-'+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].EditValue; //����� ���������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[6, 7] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('DATEPROTOKOL').Index].EditValue; //���� ���������
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('DO�TYPE').Index].EditValue=0 then
        s:='�����������'
      else
        begin
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[10, 1] :=LookUpText('DO�TYPE');  //������������ ���������
          s :='�'+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMETIKETKI').Index].EditValue;  //����� ���������
          s:=s+' �� '+DateToStr(cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ETIKETKADATE').Index].EditValue); //���� ������ ��������
          s:=s+', ����� ����������: '+LookUpText('SPETSIALIST'); //����������
          s:=s+', '+LookUpText('UPRSHNADZ'); //����. ���������� �/� ����
        end;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[11, 1]:=s;
      //�����������
      kl:=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].EditValue;
      FrmMainWindow.ADODataSet2.Active:=false;
      FrmMainWindow.ADODataSet2.CommandText:='SELECT ABBR, ADRESSCLIENT FROM CLIENTS LEFT JOIN ABBREVIATURA ON CLIENTS.ABBREVCLIENT = ABBREVIATURA.IDABBR WHERE (((CLIENTS.IDCLIENT)='+IntToStr(cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].EditValue)+'))';
      FrmMainWindow.ADODataSet2.Active:=true;
      kl:=FrmMainWindow.ADODataSet2.FieldByName('ABBR').AsString+' '+LookUpText('ORGANIZATION'); //����.�����������
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[13, 1]:=s+', '+kl; // ����������� ��� ������
      s:=kl+', '+FrmMainWindow.ADODataSet2.FieldByName('ADRESSCLIENT').AsString;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[22, 1] :=s; //����������� � �������
      //���� ������ �������������
      s:='';
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VUPRSHNADZ').Index].EditValue=true then
        s := '���������� ����������� ������ �� ������������� � ��������������� ������� �� �����-���������� � ������������� �������';
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VCLIENT').Index].EditValue=true then
        if s='' then s:=kl else s:=s+', '+kl;
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VARCHIV').Index].EditValue=true then
        if s='' then s:='� ����� ���'; // else s:=s+', � ����� ���'; // ��������� ������ 29.10.2008
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[8, 1]:=s;
      //������
      s:='';
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMZAYAVKA').Index].EditValue<>'0' then
        begin
          s:='������ �'+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NUMZAYAVKA').Index].EditValue; //����� ������ �������
          s:=s+' �� '+DateToStr(cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('DATEZAYAVKA').Index].EditValue); //���� ������
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[13, 1]:=s;
        end
      else
        begin
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[13, 1]:='';
          //ExcelApp.WorkBooks[1].WorkSheets[1].Cells[22, 1] :=''; //����������� � �������
        end;
      //�������������
      s:='';
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VIDTRANSPORT').Index].EditValue<>'0' then s:=' '+LookUpText('VIDTRANSPORT');
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NAZVTRANSPORT').Index].EditValue<>'��� ������' then s:=s+': '+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('NAZVTRANSPORT').Index].EditValue;
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].EditValue='0' then
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[20, 6] :=LookUpText('PROISHIMP')+' '+s  //������ ���
      else
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[20, 6] := cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PROISHOTECH').Index].EditValue+' '+s;
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KOLOBR').Index].EditValue<>'0' then
        begin
           if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('SHTUKVOBR').Index].EditValue<>'0' then
             ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 5] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KOLOBR').Index].EditValue + ' (' +
               cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('SHTUKVOBR').Index].EditValue+' ��. � �������)'
           else
             ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 5] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KOLOBR').Index].EditValue; //���-�� ���
        end
      else
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 5] :=''; //���-�� ���
      s:='';
      s :=LookUpText('NAMEOBRAZTSA'); //������������ �������
      s:=s+' '+LookUpText('NAMEPATRY'); //������������ ������
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('SORT').Index].EditValue<>'���' then s :=s+', ����: '+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('SORT').Index].EditValue; //����, ���� ����
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[17, 1] :=s;

      try s:=' '+LookUpText('EDVES'); except s:=' '; end;
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VES').Index].EditValue<>'0' then
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[18, 1] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VES').Index].EditValue+s //���
      else
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[18, 1] :='';
      try s:=' '+LookUpText('EDIZM'); except s:=' '; end;
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KOL').Index].EditValue<>'0' then
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[18, 4] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('KOL').Index].EditValue+s //����
      else
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[18, 4] :='';
      s:='';
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PUPRSHNADZ').Index].EditValue=true then s:='���������� ����������� ������ �� ������������� � ��������������� ������� �� �����-���������� � ������������� �������';
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PREFCENTR').Index].EditValue=true then
        if s<>'' then
          s:=s+', ��� "������������� ����������� ����� ����������� ������ �� ������������� � ��������������� �������"'
        else
          s:=s+'��� "������������� ����������� ����� ����������� ������ �� ������������� � ��������������� �������"';
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PCLIENT').Index].EditValue=true then
        if s<>'' then
          s:=s+', '+kl
        else
          s:=kl;
      ExcelApp.WorkBooks[1].WorkSheets[1].Cells[15, 1] :=s;     //----------------------------------------------------

      //���� ����������� ��������
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('DO�TYPE').Index].EditValue='0' then
        begin
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[10, 1] :='�����������';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[11, 1] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[11, 2] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[11, 3] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[11, 4] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[11, 6] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[12, 1] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[12, 2] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[12, 3] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[13, 1] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[13, 3] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[13, 4] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[13, 5] :='';
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[13, 7] :='';
        end;
      //����� ����������
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PUNKTNNO').Index].EditValue<>'0' then
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[21, 4] :=cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PUNKTNAZNACH').Index].EditValue
      else
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[21, 4] :='';
      ButtonLargeFont.Enabled:=true;
      ButtonSmallFont.Enabled:=true;
      ExcelApp.DisplayFullScreen:=true;
      ExcelApp.DisplayFullScreen:=false;
      //���������� ���������
      FrmMainWindow.ADODataSet2.Active:=false;
      s:='SELECT NAMEOBJLAT, NAMEOBJRUS, SOSTOYANIE, PRIMECH FROM EXP_KAROBJ WHERE (((IDREGISTRATION)='+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue+'))';
      s:=s+' union all SELECT NAMEOBJLAT, NAMEOBJRUS, SOSTOYANIE, PRIMECH FROM EXP_NEKAROTSRF WHERE (((IDREGISTRATION)='+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue+'))';
      s:=s+' union all SELECT NAMEOBJLAT, NAMEOBJRUS, SOSTOYANIE, PRIMECH FROM EXP_NEKARPROCH WHERE (((IDREGISTRATION)='+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue+'))';
      FrmMainWindow.ADODataSet2.CommandText:=s;
      FrmMainWindow.ADODataSet2.Active:=true;
      if FrmMainWindow.ADODataSet2.RecordCount>0 then
        for i:=0 to FrmMainWindow.ADODataSet2.RecordCount do ExcelApp.WorkBooks[1].WorkSheets[1].Cells[26, 1].EntireRow.Insert;
      //������ �����������
      FrmMainWindow.ADODataSet2.Active:=false;
      s:='SELECT NAMEOBJLAT, NAMEOBJRUS, SOSTOYANIE, PRIMECH FROM EXP_KAROBJ WHERE (((IDREGISTRATION)='+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue+'))';
      FrmMainWindow.ADODataSet2.CommandText:=s;
      FrmMainWindow.ADODataSet2.Active:=true;
      FrmMainWindow.ADODataSet2.First;
      k:=FrmMainWindow.ADODataSet2.RecordCount;
      if k<>0 then
        begin
         ExcelApp.WorkBooks[1].WorkSheets[1].Cells[25, 1] :='������������ ����������� �������:';
         for i:=1 to FrmMainWindow.ADODataSet2.RecordCount do
          begin
           if Length(FrmMainWindow.ADODataSet2.FieldByName('SOSTOYANIE').AsString)>1 then c:='('+FrmMainWindow.ADODataSet2.FieldByName('SOSTOYANIE').AsString+')' else c:='';
           s:=FrmMainWindow.ADODataSet2.FieldByName('NAMEOBJLAT').AsString+' '+FrmMainWindow.ADODataSet2.FieldByName('NAMEOBJRUS').AsString+' '+c;
           ExcelApp.WorkBooks[1].WorkSheets[1].Cells[i+25, 1] :=' '+IntToStr(i)+'. '+s;
           FrmMainWindow.ADODataSet2.Next;
          end;
         end
      else
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[25, 1] :='����������� ������� - �� ����������.';
      //������ �������������, ������������� � ��
      k:=k+1;
      FrmMainWindow.ADODataSet2.Active:=false;
      s:='SELECT NAMEOBJLAT, NAMEOBJRUS, SOSTOYANIE, PRIMECH FROM EXP_NEKAROTSRF WHERE (((IDREGISTRATION)='+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue+'))';
      FrmMainWindow.ADODataSet2.CommandText:=s;
      FrmMainWindow.ADODataSet2.Active:=true;
      FrmMainWindow.ADODataSet2.First;
      if FrmMainWindow.ADODataSet2.RecordCount<>0 then
        begin
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[25+k, 1] :='������������ ������������� �������, ������������� � ��:';
          for i:=1 to FrmMainWindow.ADODataSet2.RecordCount do
            begin
             if Length(FrmMainWindow.ADODataSet2.FieldByName('SOSTOYANIE').AsString)>1 then c:='('+FrmMainWindow.ADODataSet2.FieldByName('SOSTOYANIE').AsString+')' else c:='';
             s:=FrmMainWindow.ADODataSet2.FieldByName('NAMEOBJLAT').AsString+' '+FrmMainWindow.ADODataSet2.FieldByName('NAMEOBJRUS').AsString+' '+c;
             ExcelApp.WorkBooks[1].WorkSheets[1].Cells[i+25+k, 1] :=' '+IntToStr(i)+'. '+s;
             FrmMainWindow.ADODataSet2.Next;
            end;
          k:=k+FrmMainWindow.ADODataSet2.RecordCount;
        end
      else
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[25+k, 1] :='������������� �������, ������������� � �� - �� ����������.';
      //������ �������������, ������
      k:=k+1;
      FrmMainWindow.ADODataSet2.Active:=false;
      s:='SELECT NAMEOBJLAT, NAMEOBJRUS, SOSTOYANIE, PRIMECH FROM EXP_NEKARPROCH WHERE (((IDREGISTRATION)='+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue+'))';
      FrmMainWindow.ADODataSet2.CommandText:=s;
      FrmMainWindow.ADODataSet2.Active:=true;
      FrmMainWindow.ADODataSet2.First;
      if FrmMainWindow.ADODataSet2.RecordCount<>0 then
        begin
          ExcelApp.WorkBooks[1].WorkSheets[1].Cells[25+k, 1] :='������������ ������������� �������, ������:';
          for i:=1 to FrmMainWindow.ADODataSet2.RecordCount do
            begin
              if Length(FrmMainWindow.ADODataSet2.FieldByName('SOSTOYANIE').AsString)>1 then c:='('+FrmMainWindow.ADODataSet2.FieldByName('SOSTOYANIE').AsString+')' else c:='';
              s:=FrmMainWindow.ADODataSet2.FieldByName('NAMEOBJLAT').AsString+' '+FrmMainWindow.ADODataSet2.FieldByName('NAMEOBJRUS').AsString+' '+c;
              ExcelApp.WorkBooks[1].WorkSheets[1].Cells[i+25+k, 1] :=' '+IntToStr(i)+'. '+s;
              FrmMainWindow.ADODataSet2.Next;
            end;
          k:=k+FrmMainWindow.ADODataSet2.RecordCount;
        end
      else
        ExcelApp.WorkBooks[1].WorkSheets[1].Cells[25+k, 1] :='������������� �������, ������ - �� ����������.';
      //���������� ��������� �����������
      FrmMainWindow.ADODataSet2.Active:=false;
      s:='SELECT SPETSIALISTS_LAB.NAMESPETSIALIST, SPETSIALISTS_LAB.DOLZHNSPETSIALIST FROM EXPERTIZE_REESTR INNER JOIN SPETSIALISTS_LAB ON EXPERTIZE_REESTR.SPETSIALIST = SPETSIALISTS_LAB.IDSPETSIALIST WHERE (((EXPERTIZE_REESTR.NUMPROTOKOL)=';
      s:=s+cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue+')) GROUP BY SPETSIALISTS_LAB.NAMESPETSIALIST, SPETSIALISTS_LAB.DOLZHNSPETSIALIST';
      FrmMainWindow.ADODataSet2.CommandText:=s;
      FrmMainWindow.ADODataSet2.Active:=true;
      FrmMainWindow.ADODataSet2.First;
      if FrmMainWindow.ADODataSet2.RecordCount<>0 then
        begin
          for i:=0 to FrmMainWindow.ADODataSet2.RecordCount-2 do
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[25+k+4, 1].EntireRow.Insert;
          for i:=1 to FrmMainWindow.ADODataSet2.RecordCount do
            begin
              s:=FrmMainWindow.ADODataSet2.FieldByName('DOLZHNSPETSIALIST').AsString+' '+FrmMainWindow.ADODataSet2.FieldByName('NAMESPETSIALIST').AsString;
              ExcelApp.WorkBooks[1].WorkSheets[1].Cells[i+25+k+2, 1] :=' '+IntToStr(i)+'. '+s;
              FrmMainWindow.ADODataSet2.Next;
            end;
        end;
      //��� ������������� - ��������� ������ ����������: �����������, ������������� ...
      if cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].EditValue='1' then
        begin
          FrmMainWindow.ADODataSet2.Active:=false;
          FrmMainWindow.ADODataSet2.CommandText:='SELECT * FROM OTECHMEMOSORT WHERE (((OTECHMEMOSORT.IDREGISTRATION)='+IntToStr(StrToInt(cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].EditValue)+1)+'))';
          FrmMainWindow.ADODataSet2.Active:=true;
          if FrmMainWindow.ADODataSet2.RecordCount>0 then
            begin
              ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 1].EntireRow.Insert;
              ExcelApp.WorkBooks[1].WorkSheets[1].Range[ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 1], ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 10]].Merge;
              ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 1].EntireRow.Insert;
              ExcelApp.WorkBooks[1].WorkSheets[1].Range[ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 1], ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 10]].Merge;
            end;
          if copy(FrmMainWindow.ADODataSet2.FieldByName('SORTTMPROCH').AsString,1,10)<>'��� ������' then ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 1] :=FrmMainWindow.ADODataSet2.FieldByName('SORTTMPROCH').AsString;
          //ExcelApp.WorkBooks[1].WorkSheets[1].Cells[20, 1] :=FrmMainWindow.ADODataSet2.FieldByName('PROISH').AsString;
          //ExcelApp.WorkBooks[1].WorkSheets[1].Cells[i+25, 1]
        end;
      //�������� ���� �� ���������
      //ExcelApp.WorkBooks[1].WorkSheets[1].protect('344',true,true,false);
end;

procedure TFrmGrid1.ButtonExpandClick(Sender: TObject);
begin
   cxGrid1DBTableView1.ViewData.Expand(True);
end;

procedure TFrmGrid1.ButtonCollapsClick(Sender: TObject);
begin
   cxGrid1DBTableView1.ViewData.Collapse(True);
end;

procedure TFrmGrid1.cxGrid1DBTableView1CellDblClick(
  Sender: TcxCustomGridTableView;
  ACellViewInfo: TcxGridTableDataCellViewInfo; AButton: TMouseButton;
  AShift: TShiftState; var AHandled: Boolean);
var
  i, u     : integer;
  SQLstr : string;
begin
  if (cxGrid1DBTableView1.Controller.FocusedColumn.Caption='������ ������������� ��� ���.���.') and (cxGrid1DBTableView1.DataController.DataSource.DataSet.FieldByName('ISIMPORT').AsString='0') then
    begin
        FrmProishImpObr.ADODataSet2.Active:=true;
        FrmProishImpObr.cxGrid1DBTableView1.DataController.CreateAllItems(true);
        with FrmProishImpObr.cxGrid1DBTableView1 do
        begin
          Columns[FrmProishImpObr.cxGrid1DBTableView1.GetColumnByFieldName('IDPROISH').Index].Caption:='��� ������';
          Columns[FrmProishImpObr.cxGrid1DBTableView1.GetColumnByFieldName('NAMEPROISH').Index].Caption:='������';
          Columns[FrmProishImpObr.cxGrid1DBTableView1.GetColumnByFieldName('NAMEPROISH').Index].Width:=140;
        end;
        FrmProishImpObr.Showmodal;
        if PerStart='' then exit;
        Screen.Cursor:=crHourGlass;
        cxGrid1DBTableView1.DataController.SaveBookmark;
        FrmProishImpObr.ADOCommand1.CommandText:='UPDATE REGISTRATION SET REGISTRATION.PROISHOZHD = '+PerStart+' WHERE (((REGISTRATION.IDREGISTRATION)='+cxGrid1DBTableView1.DataController.DataSource.DataSet.FieldByName('IDREGISTRATION').AsString+'));';
        FrmProishImpObr.ADOCommand1.Execute;
        FrmMainWindow.ADODataSet1.Active:=false;
        FrmMainWindow.ADODataSet1.Active:=true;
        //�������� �������
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VUPRSHNADZ').Index].Visible:=false;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VCLIENT').Index].Visible:=false;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VARCHIV').Index].Visible:=false;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PUPRSHNADZ').Index].Visible:=false;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PCLIENT').Index].Visible:=false;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PREFCENTR').Index].Visible:=false;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('EXPCLOSE').Index].Visible:=false;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('OTKAZKL').Index].Visible:=false;
        //������ ���������� ����� ��������� ��������
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VUPRSHNADZ').Index].Hidden:=true;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VCLIENT').Index].Hidden:=true;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VARCHIV').Index].Hidden:=true;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PUPRSHNADZ').Index].Hidden:=true;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PCLIENT').Index].Hidden:=true;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PREFCENTR').Index].Hidden:=true;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('EXPCLOSE').Index].Hidden:=true;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('OTKAZKL').Index].Hidden:=true;
        cxGrid1DBTableView1.DataController.GotoBookmark;
        Screen.Cursor:=crDefault;
    end;
  if (cxGrid1DBTableView1.Controller.FocusedColumn.Caption='��������� ������������� ��� ����.���.') and (cxGrid1DBTableView1.DataController.DataSource.DataSet.FieldByName('ISIMPORT').AsString='1') then
    begin
      if FrmProishObrOtech= NIL then FrmProishObrOtech := TFrmProishObrOtech.Create(owner);
      Screen.Cursor:=crHourGlass;
      FrmProishObrOtech.ADODataSet4.Active:=false;
      FrmProishObrOtech.cxGrid1DBTableView1.ClearItems;
      FrmProishObrOtech.ADODataSet4.Active:=true;
      with FrmProishObrOtech.cxGrid1DBTableView1.DataController do
        CreateAllItems;
      with FrmProishObrOtech.cxGrid1DBTableView1 do
        begin
          Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('IDNASPUNKT').Index].Visible:=false;
          Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('IDNASPUNKT').Index].Caption:='��� ���������� ������';
          Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('OBLNAME').Index].Caption:='�������';
          Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('RAYONNAME').Index].Caption:='�����';
          Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Caption:='��� ���������� ������';
          Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTNAME').Index].Caption:='�������� ���������� ������';
        end;
      FrmProishObrOtech.cxGrid1DBTableView1.Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].PropertiesClass:=TcxImageComboBoxProperties;
      with FrmProishObrOtech.cxGrid1DBTableView1 do
        begin
          //TcxImageComboBoxProperties(FrmProishObrOtech.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Images:=FrmGrid1.ImageList2;
          TcxImageComboBoxProperties(FrmProishObrOtech.cxGrid1DBTableView1.Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items.Add;
          TcxImageComboBoxProperties(FrmProishObrOtech.cxGrid1DBTableView1.Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items.Add;
          TcxImageComboBoxProperties(FrmProishObrOtech.cxGrid1DBTableView1.Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items.Add;
          TcxImageComboBoxProperties(FrmProishObrOtech.cxGrid1DBTableView1.Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items.Add;
          TcxImageComboBoxProperties(FrmProishObrOtech.cxGrid1DBTableView1.Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items.Add;
          TcxImageComboBoxProperties(Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items[0].Description:='�������';
          TcxImageComboBoxProperties(Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items[0].Value:='�������';
          TcxImageComboBoxProperties(Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items[1].Description:='������';
          TcxImageComboBoxProperties(Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items[1].Value:='������';
          TcxImageComboBoxProperties(Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items[2].Description:='�����';
          TcxImageComboBoxProperties(Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items[2].Value:='�����';
          TcxImageComboBoxProperties(Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items[3].Description:='�������� �����';
          TcxImageComboBoxProperties(Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items[3].Value:='�������� �����';
          TcxImageComboBoxProperties(Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items[3].Description:='�����';
          TcxImageComboBoxProperties(Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items[3].Value:='�����';
        end;
      for i:=0 to FrmProishObrOtech.cxGrid1DBTableView1.ColumnCount-1 do FrmProishObrOtech.cxGrid1DBTableView1.Columns[i].Width:=120;
      Screen.Cursor:=crDefault;
      FrmProishObrOtech.ShowModal;
      if PerStart='' then exit;
        Screen.Cursor:=crHourGlass;
        cxGrid1DBTableView1.DataController.SaveBookmark;
        FrmProishImpObr.ADOCommand1.CommandText:='UPDATE OTECHMEMOSORT SET OTECHMEMOSORT.PROISH = "'+PerStart+'" WHERE (((OTECHMEMOSORT.IDREGISTRATION)='+cxGrid1DBTableView1.DataController.DataSource.DataSet.FieldByName('IDREGISTRATION').AsString+'+1));';
        FrmProishImpObr.ADOCommand1.Execute;
        FrmMainWindow.ADODataSet1.Active:=false;
        FrmMainWindow.ADODataSet1.Active:=true;
        //�������� �������
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VUPRSHNADZ').Index].Visible:=false;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VCLIENT').Index].Visible:=false;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VARCHIV').Index].Visible:=false;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PUPRSHNADZ').Index].Visible:=false;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PCLIENT').Index].Visible:=false;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PREFCENTR').Index].Visible:=false;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('EXPCLOSE').Index].Visible:=false;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('OTKAZKL').Index].Visible:=false;
        //������ ���������� ����� ��������� ��������
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VUPRSHNADZ').Index].Hidden:=true;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VCLIENT').Index].Hidden:=true;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('VARCHIV').Index].Hidden:=true;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PUPRSHNADZ').Index].Hidden:=true;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PCLIENT').Index].Hidden:=true;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('PREFCENTR').Index].Hidden:=true;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('EXPCLOSE').Index].Hidden:=true;
        cxGrid1DBTableView1.Columns[cxGrid1DBTableView1.GetColumnByFieldName('OTKAZKL').Index].Hidden:=true;
        cxGrid1DBTableView1.DataController.GotoBookmark;
        Screen.Cursor:=crDefault;
    end;
    if ChangeStatus=true then
       begin
          if PMMessageDialog('�� �������� ���������� � ����� ������ � �����������.'+Chr(13)+'����������� ������?', '��������:', mtConfirmation, mbOKCancel, ['��', '������']) = mrOk then
             begin
                {
                UpdStr - ������ ������ ������ � ������� ������
                UpdFld - ���� � ������� ������, ������� ��������
                UpdSpr - ������� �����������
                UpdKey - ���� ���� � ������� �����������
                }
                u:=FrmMainWindow.ADODataSet1.FieldByName(UpdKey).AsInteger;
                //� ������� �����������
                if copy(UpdStr,1,19)='UPDATE REGISTRATION' then
                  begin
                    SQLstr:=UpdStr + IntToStr(u)+' WHERE '+UpdFld+'='+IntToStr(OldValChange);
                    ADOCommand1.CommandText:=SQLstr;
                    ADOCommand1.Execute;
                    //� ����� �����������
                    FrmMainWindow.ADODataSet1.Active:=false;
                    SQLstr:='UPDATE '+UpdSpr+' SET IDCHANGE = '+IntToStr(u)+' WHERE ((('+UpdKey+')='+IntToStr(OldValChange)+'))';
                    ADOCommand1.CommandText:=SQLstr;
                    showmessage(SQLstr);
                    ADOCommand1.Execute;
                  end;
                //� ������� ��������� ������ update �� �������� - � ����� ������� ���������
                if copy(UpdStr,1,21)='UPDATE EXPERTIZE_DATA' then
                  begin
                    SQLstr:=UpdStr + IntToStr(u)+' WHERE (('+UpdFld+'='+IntToStr(OldValChange)+') AND (TYPEOBJ='+IntToStr(UpdKarObj)+'))';
                    ADOCommand1.CommandText:=SQLstr;
                    ADOCommand1.Execute; 
                    //� ����� �����������
                    FrmMainWindow.ADODataSet1.Active:=false;
                    SQLstr:='UPDATE '+UpdSpr+' SET IDCHANGE = '+IntToStr(u)+' WHERE ((('+UpdKey+')='+IntToStr(OldValChange)+'))';
                    ADOCommand1.CommandText:=SQLstr;
                    //showmessage(SQLstr);
                    ADOCommand1.Execute;
                  end;
                FrmMainWindow.ADODataSet1.Active:=true;
             end;
          ChangeStatus:=false;
       end;
end;

procedure TFrmGrid1.ButtonDeleteProtokolClick(Sender: TObject);
begin
   if PMMessageDialog('��� �������� ���������, ��� ������, ��������� � ���, ���������!'+Chr(13)+'������� ��������?', '��������:', mtWarning, mbOKCancel, ['�������', '������']) = mrOk then
      begin
         FrmProishImpObr.ADOCommand1.CommandText:='DELETE EXPERTIZE_DATA.IDROW FROM EXPERTIZE_DATA WHERE (((EXPERTIZE_DATA.IDROW)='+FrmMainWindow.ADODataSet1.FieldByName('IDROW').AsString+'))';
         FrmProishImpObr.ADOCommand1.Execute;
         FrmMainWindow.ADODataSet1.Active:=false;
         FrmMainWindow.ADODataSet1.Active:=true;
      end;
end;

procedure TFrmGrid1.ButtonSwapClick(Sender: TObject);
begin
   if UpdStr='' then
      begin
         ButtonSwap.Visible:=false;
         exit;
      end;
   if PMMessageDialog('�� ������� � ����� ������ ������������.'+Chr(13)+'��������� �� ������, ������� ������� ������ � �������� ������� ����.', '��������:', mtWarning, mbOKCancel, ['�������.', '������']) = mrOk then
      begin
         ChangeStatus:=true;
         OldValChange:=FrmMainWindow.ADODataSet1.FieldByName(UpdKey).AsInteger;
         cxGrid1DBTableView1.OptionsData.Editing:=false;  // ��������� ����������� ������� � �������� ���������
      end;
end;

procedure TFrmGrid1.ButtonCopyClick(Sender: TObject);
var
   Layout: array[0.. KL_NAMELENGTH] of char;
begin
   //������������ � ������ ���������
   LoadKeyboardLayout(StrCopy(Layout,'00000419'),KLF_ACTIVATE);
   if cxGrid1.ActiveView=cxGrid1DBBandedTableView1 then
     try
       cxGrid1DBBandedTableView1.CopyToClipboard(false);
     except
       showmessage('����������� �� �����������: ��� ���������� ����� ��� ������ ������ ������.');
     end
   else  
     try
       cxGrid1DBTableView1.CopyToClipboard(false);
     except
       showmessage('����������� �� �����������: ��� ���������� ����� ��� ������ ������ ������.');
     end;
end;

end.
