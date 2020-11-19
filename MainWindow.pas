unit MainWindow;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, AdvGlowButton, AdvToolBar, cxStyles, AdvToolBarStylers, DB,
  ADODB, ExtCtrls, ComObj, cxGraphics, cxControls, dxStatusBar, cxGridPopupMenuConsts,
  cxContainer, cxEdit, cxTextEdit, cxMaskEdit, cxDropDownEdit, cxGrid,  cxCalendar,
  cxLookupEdit, cxDBLookupEdit, cxDBLookupComboBox, cxImageComboBox, Menus,
  AdvPreviewMenu, AdvPreviewMenuStylers, AdvShapeButton, Registry,
  StdCtrls, AdvSmoothCalendar, AdvSmoothLabel, AdvSmoothPanel, DateUtils,
  AdvWiiProgressBar;

type
  TFrmMainWindow = class(TForm)
    AdvToolBarPager1: TAdvToolBarPager;
    AdvPage1: TAdvPage;
    AdvPage2: TAdvPage;
    AdvToolBar1: TAdvToolBar;
    ButtonCreateProtokolImport: TAdvGlowButton;
    StyleRepository: TcxStyleRepository;
    AdvToolBarFantasyStyler1: TAdvToolBarFantasyStyler;
    ButtonFindProtokolImport: TAdvGlowButton;
    AdvPage3: TAdvPage;
    AdvToolBar3: TAdvToolBar;
    ButtonShowClients: TAdvGlowButton;
    AdvToolBar4: TAdvToolBar;
    ButtonExpSorny: TAdvGlowButton;
    ButtonExpFind: TAdvGlowButton;
    AdvToolBar6: TAdvToolBar;
    AdvToolBar5: TAdvToolBar;
    AdvGlowButton19: TAdvGlowButton;
    AdvGlowButton20: TAdvGlowButton;
    AdvToolBar2: TAdvToolBar;
    AdvGlowButton3: TAdvGlowButton;
    AdvGlowButton4: TAdvGlowButton;
    ADOConnection1: TADOConnection;
    ADODataSet1: TADODataSet;
    DataSource1: TDataSource;
    ButtonExpNematod: TAdvGlowButton;
    ButtonExpVirus: TAdvGlowButton;
    ButtonExpBakter: TAdvGlowButton;
    ButtonExpGrib: TAdvGlowButton;
    ButtonExpVreditel: TAdvGlowButton;
    AdvToolBarSeparator1: TAdvToolBarSeparator;
    AdvPage4: TAdvPage;
    AdvToolBar7: TAdvToolBar;
    ButtonNewReg: TAdvGlowButton;
    ButtonComputers: TAdvGlowButton;
    AdvToolBar8: TAdvToolBar;
    ButtonProtokolTemplate: TAdvGlowButton;
    ButtonSvvoTemplate: TAdvGlowButton;
    AdvToolBar9: TAdvToolBar;
    AdvToolBar10: TAdvToolBar;
    ButtonShowAbbreviatura: TAdvGlowButton;
    ButtonShowUprNadzora: TAdvGlowButton;
    ButtonShowSpetsialist: TAdvGlowButton;
    ButtonShowPunkts: TAdvGlowButton;
    ButtonShowObrazets: TAdvGlowButton;
    ButtonShowProishozhd: TAdvGlowButton;
    ButtonShowPartii: TAdvGlowButton;
    ButtonShowKlass: TAdvGlowButton;
    ButtonShowTransport: TAdvGlowButton;
    ButtonShowVes: TAdvGlowButton;
    ButtonShowEdIzm: TAdvGlowButton;
    AdvToolBarSeparator3: TAdvToolBarSeparator;
    AdvToolBarSeparator2: TAdvToolBarSeparator;
    ButtonShowKarObj: TAdvGlowButton;
    ButtonNekarObjProchee: TAdvGlowButton;
    ButtonNekarObjOtsutRF: TAdvGlowButton;
    AdvToolBar11: TAdvToolBar;
    ButtonShowSpetsLab: TAdvGlowButton;
    AdvToolBar12: TAdvToolBar;
    ButtonAbout: TAdvGlowButton;
    ButtonShowDocVids: TAdvGlowButton;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    PopupMenu2: TPopupMenu;
    N21: TMenuItem;
    N22: TMenuItem;
    N23: TMenuItem;
    N24: TMenuItem;
    N25: TMenuItem;
    N26: TMenuItem;
    AdvToolBar13: TAdvToolBar;
    AdvGlowButton2: TAdvGlowButton;
    ADODataSet2: TADODataSet;
    DataSource2: TDataSource;
    ButtonShowExcel: TAdvGlowButton;
    ButtonSvodOtch: TAdvGlowButton;
    ButtonSvodTemplate: TAdvGlowButton;
    ButtonSvvo2Template: TAdvGlowButton;
    ButtonAktTemplate: TAdvGlowButton;
    AdvGlowButton5: TAdvGlowButton;
    Timer1: TTimer;
    ButtonZhurnal: TAdvGlowButton;
    ButtonSvvo3Template: TAdvGlowButton;
    ADODataSetLockRegistration: TADODataSet;
    ADOCommandLockRegistration: TADOCommand;
    ButtonUnblockRegProtokol: TAdvGlowButton;
    PreviewMenu1: TAdvPreviewMenu;
    AdvPreviewMenuOfficeStyler1: TAdvPreviewMenuOfficeStyler;
    AdvShapeButton2: TAdvShapeButton;
    dxStatusBar1: TdxStatusBar;
    Panel1: TPanel;
    ADODataSet_Statistics: TADODataSet;
    Memo_Statistics: TMemo;
    Timer2: TTimer;
    WiiProgressBar1: TAdvWiiProgressBar;
    procedure ButtonFindProtokolImportClick(Sender: TObject);
    procedure ButtonCreateProtokolImportClick(Sender: TObject);
    procedure ToForm(Sender: TForm);
    procedure ExpertMaster(hintmaster: string);
    procedure ButtonShowClientsClick(Sender: TObject);
    procedure ButtonExpVreditelClick(Sender: TObject);
    procedure ButtonExpGribClick(Sender: TObject);
    procedure ButtonExpBakterClick(Sender: TObject);
    procedure ButtonExpVirusClick(Sender: TObject);
    procedure ButtonExpNematodClick(Sender: TObject);
    procedure ButtonExpSornyClick(Sender: TObject);
    procedure ButtonShowKarObjClick(Sender: TObject);
    procedure ButtonShowSpetsialistClick(Sender: TObject);
    procedure ButtonShowUprNadzoraClick(Sender: TObject);
    procedure ButtonShowObrazetsClick(Sender: TObject);
    procedure ButtonShowPartiiClick(Sender: TObject);
    procedure ButtonShowKlassClick(Sender: TObject);
    procedure ButtonShowProishozhdClick(Sender: TObject);
    procedure ButtonShowEdIzmClick(Sender: TObject);
    procedure ButtonShowVesClick(Sender: TObject);
    procedure ButtonShowTransportClick(Sender: TObject);
    procedure ButtonNekarObjProcheeClick(Sender: TObject);
    procedure ButtonComputersClick(Sender: TObject);
    procedure ButtonProtokolTemplateClick(Sender: TObject);
    procedure ButtonSvvoTemplateClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ButtonShowSpetsLabClick(Sender: TObject);
    procedure ButtonShowAbbreviaturaClick(Sender: TObject);
    procedure ButtonShowPunktsClick(Sender: TObject);
    procedure ButtonShowDocVidsClick(Sender: TObject);
    procedure NekarObjOtsRF(typeobj: string);
    procedure NekarObjProchie(typeobj: string);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure ADODataSet1NewRecord(DataSet: TDataSet);
    procedure SelSpetsialistForExp(sender: TObject);
    procedure ButtonExpFindClick(Sender: TObject);
    procedure AdvGlowButton2Click(Sender: TObject);
    procedure ButtonShowExcelClick(Sender: TObject);
    procedure ButtonAboutClick(Sender: TObject);
    procedure ButtonNewRegClick(Sender: TObject);
    procedure N21Click(Sender: TObject);
    procedure N22Click(Sender: TObject);
    procedure N26Click(Sender: TObject);
    procedure N23Click(Sender: TObject);
    procedure N24Click(Sender: TObject);
    procedure N25Click(Sender: TObject);
    procedure ButtonSvodTemplateClick(Sender: TObject);
    procedure ButtonSvodOtchClick(Sender: TObject);
    procedure ButtonSvvo2TemplateClick(Sender: TObject);
    procedure ButtonAktTemplateClick(Sender: TObject);
    procedure AdvGlowButton5Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure ButtonZhurnalClick(Sender: TObject);
    procedure ButtonSvvo3TemplateClick(Sender: TObject);
    procedure ButtonUnblockRegProtokolClick(Sender: TObject);
    procedure AdvShapeButton2Click(Sender: TObject);
    procedure PreviewMenu1Buttons0Click(Sender: TObject;
      Button: TButtonCollectionItem);
    procedure PreviewMenu1MenuItems0Click(Sender: TObject);
    procedure PreviewMenu1MenuItems2Click(Sender: TObject);
    procedure PreviewMenu1MenuItems1Click(Sender: TObject);
    procedure ButtonNekarObjOtsutRFClick(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure WiiProgressBar1Click(Sender: TObject);
  private
    procedure ValidateAccessDB;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmMainWindow                                  : TFrmMainWindow;
  DateLastDog, DateTekPerStart, DateTekPerFinish,
  KeyFieldDet, ExpTypeName, PerStart, PerFinish,
  DefaultVal, UpdStr, UpdKey, UpdFld, UpdSpr     : string;
  TblRezh, fc, UpdKarObj,
  TypeNumeration, LastNumProtForManual,
  LastNumProtForAuto                             : integer;
  FullDostupBool, ShowNotFoundObj, Oeape, MUReg,
  ViewStatistic                                  : boolean;
  ExcelApp, Workbook                             : Variant;

implementation

uses Grid1, Dialog1, ComCtrls, Dialog2, cxGridTableView, Period, PeriodExp,
  PeriodOtch, About, ToolExcel, Options, NumProtocols, Logo, Calendar;

{$R *.dfm}

function GetLoginName: string;
var
  buffer: array[0..255] of char;
  size: dword;
begin
  size := 256;
  if GetUserName(buffer, size) then
    Result := buffer
  else
    Result := ''
end;

function GetComputerNetName: string;
var
  buffer: array[0..255] of char;
  size: dword;
begin
  size := 256;
  if GetComputerName(buffer, size) then
    Result := buffer
  else
    Result := ''
end;

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

procedure TFrmMainWindow.ButtonFindProtokolImportClick(Sender: TObject);
var
  i:integer;
  SQLstr: string;
begin
  FrmPeriod.Showmodal;
  if PerStart='' then exit;
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if FrmDialog1= NIL then FrmDialog1 := TFrmDialog1.Create(owner);
  if ADOConnection1.Connected=false then ADOConnection1.Connected:=true;
  //обновление источников данных
  FrmDialog1.ADODataSet1.Active:=false;
  FrmDialog1.ADODataSet2.Active:=false;
  FrmDialog1.ADODataSet3.Active:=false;
  FrmDialog1.ADODataSet4.Active:=false;
  FrmDialog1.ADODataSet5.Active:=false;
  FrmDialog1.ADODataSet6.Active:=false;
  FrmDialog1.ADODataSet7.Active:=false;
  FrmDialog1.ADODataSet8.Active:=false;
  FrmDialog1.ADODataSet9.Active:=false;
  FrmDialog1.ADODataSet10.Active:=false;
  FrmDialog1.ADODataSet11.Active:=false;
  FrmDialog1.ADODataSet13.Active:=false;
  FrmDialog1.ADODataSet1.Active:=true;
  FrmDialog1.ADODataSet2.Active:=true;
  FrmDialog1.ADODataSet3.Active:=true;
  FrmDialog1.ADODataSet4.Active:=true;
  FrmDialog1.ADODataSet5.Active:=true;
  FrmDialog1.ADODataSet6.Active:=true;
  FrmDialog1.ADODataSet7.Active:=true;
  FrmDialog1.ADODataSet8.Active:=true;
  FrmDialog1.ADODataSet9.Active:=true;
  FrmDialog1.ADODataSet10.Active:=true;
  FrmDialog1.ADODataSet11.Active:=true;
  FrmDialog1.ADODataSet13.Active:=true;
  ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  SQLstr:='SELECT REGISTRATION.*, IIf([REGISTRATION]![ISIMPORT]=0,[REGISTRATION]![PROISHOZHD],0) AS PROISHIMP, OTECHMEMOSORT.PROISH AS PROISHOTECH';
  SQLstr:=SQLstr+' FROM REGISTRATION LEFT JOIN OTECHMEMOSORT ON REGISTRATION.IDREGISTRATION +1 = OTECHMEMOSORT.IDREGISTRATION WHERE (((DATEPROTOKOL)>=#'+PerStart+'# And (DATEPROTOKOL)<=#'+PerFinish+'#))';
  SQLstr:=SQLstr+' ORDER BY REGISTRATION.IDREGISTRATION DESC';
  ADODataSet1.CommandText:=SQLstr;
  ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].Caption:='Номер записи';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Caption:='Признак происхождения образца';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].Caption:='Номер протокола';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DATEPROTOKOL').Index].Caption:='Дата оформления';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DOСTYPE').Index].Caption:='Тип документа';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('UPRSHNADZ').Index].Caption:='Наименование Управления Россельхознадзора';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PUNKTKARRAST').Index].Caption:='Название пункта по карантину растений';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('SPETSIALIST').Index].Caption:='Ф.И.О. специалиста, выдавшего этикетку';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NUMETIKETKI').Index].Caption:='Номер этикетки';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ETIKETKADATE').Index].Caption:='Дата выдачи этикетки';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].Caption:='Название организации (Ф.И.О. частного лица)';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NUMDOGOVOR').Index].Caption:='Номер договора';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NUMZAYAVKA').Index].Caption:='Номер заявки клиента';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DATEZAYAVKA').Index].Caption:='Дата заявки клиента';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEOBRAZTSA').Index].Caption:='Наименование образца';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('KLASSIFIKATION').Index].Caption:='Классификация';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEPATRY').Index].Caption:='Наименование партии продукции, от которой отобран образец';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('SORT').Index].Caption:='Сорт образца';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PROISHOZHD').Index].Caption:='Код происхождение имп.образца';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PROISHOZHD').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('VES').Index].Caption:='Вес';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EDVES').Index].Caption:='Ед.изм. веса';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('KOL').Index].Caption:='Количество мест';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EDIZM').Index].Caption:='Ед.изм. мест';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('KOLOBR').Index].Caption:='Количество образцов';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('SHTUKVOBR').Index].Caption:='Штук в образце';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('VIDTRANSPORT').Index].Caption:='Наименование транспортного средства';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAZVTRANSPORT').Index].Caption:='Название транспортного средства, номер, приход';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PUNKTNAZNACH').Index].Caption:='Пункт назначения (страна, регион, клиент)';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('VUPRSHNADZ').Index].Caption:=' Управлению Россельхознадзора выдано Свидетельство карантинной экспертизы ';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('VCLIENT').Index].Caption:='Клиенту выдано Свидетельство карантинной экспертизы ';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('VARCHIV').Index].Caption:='В архиве ФГУ находится Свидетельство карантинной экспертизы ';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PUPRSHNADZ').Index].Caption:='Из Управления Россельхознадзора поступили образцы';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PCLIENT').Index].Caption:='От клиента поступили образцы';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PREFCENTR').Index].Caption:='От Референтного центра поступили образцы';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EXPCLOSE').Index].Caption:='Свидетельство выдано';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTKAZKL').Index].Caption:='Отказ клиента от анализа';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PROISHIMP').Index].Caption:='Страна происхождения для имп.обр.';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PROISHOTECH').Index].Caption:='Местность происхождения для отеч.обр.';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAZNRF').Index].Caption:='Профиль назначения - Российская Федерация';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAZNEXP').Index].Caption:='Профиль назначения - экспорт';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PUNKTNNO').Index].Caption:='Адрес клиента и пункта назначения совпадают';
      //свойства столбцов
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DATEPROTOKOL').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PROISHIMP').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PROISHOTECH').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PUNKTNNO').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].PropertiesClass:=TcxImageComboBoxProperties;
      with FrmGrid1.cxGrid1DBTableView1 do
        begin
          TcxImageComboBoxProperties(FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Images:=FrmGrid1.ImageList2;
          TcxImageComboBoxProperties(FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items.Add;
          TcxImageComboBoxProperties(FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items.Add;
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items[0].Description:='Импортный';
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items[0].Value:=0;
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items[1].Description:='Отечественный';
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items[1].Value:=1;
        end;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DOСTYPE').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DOСTYPE').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource13;
          ListFieldNames := 'DOCVIDS';
          KeyFieldNames := 'IDDOCVIDS';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DOСTYPE').Index].DataBinding.FieldName := 'DOСTYPE';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('UPRSHNADZ').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('UPRSHNADZ').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource3;
          ListFieldNames := 'NAMEUPRSHNADZ';
          KeyFieldNames := 'IDUPRSHNADZ';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('UPRSHNADZ').Index].DataBinding.FieldName := 'UPRSHNADZ';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PUNKTKARRAST').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PUNKTKARRAST').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource2;
          ListFieldNames := 'NAMEKARPUNKT';
          KeyFieldNames := 'IDKARPUNKT';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PUNKTKARRAST').Index].DataBinding.FieldName := 'PUNKTKARRAST';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('SPETSIALIST').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('SPETSIALIST').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource1;
          ListFieldNames := 'NAMESPETSIALIST';
          KeyFieldNames := 'IDSPETSIALIST';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('SPETSIALIST').Index].DataBinding.FieldName := 'SPETSIALIST';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource4;
          ListFieldNames := 'NAMECLIENT';
          KeyFieldNames := 'IDCLIENT';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].DataBinding.FieldName := 'ORGANIZATION';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEOBRAZTSA').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEOBRAZTSA').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource5;
          ListFieldNames := 'NAMEOBR';
          KeyFieldNames := 'IDOBR';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEOBRAZTSA').Index].DataBinding.FieldName := 'NAMEOBRAZTSA';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('KLASSIFIKATION').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('KLASSIFIKATION').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource7;
          ListFieldNames := 'NAMEKLASS';
          KeyFieldNames := 'IDKLASS';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('KLASSIFIKATION').Index].DataBinding.FieldName := 'KLASSIFIKATION';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEPATRY').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEPATRY').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource6;
          ListFieldNames := 'NAMEPARTII';
          KeyFieldNames := 'IDPARTII';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEPATRY').Index].DataBinding.FieldName := 'NAMEPATRY';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PROISHIMP').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PROISHIMP').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource8;
          ListFieldNames := 'NAMEPROISH';
          KeyFieldNames := 'IDPROISH';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PROISHIMP').Index].DataBinding.FieldName := 'PROISHIMP';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EDVES').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EDVES').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource9;
          ListFieldNames := 'NAMEMEASURE';
          KeyFieldNames := 'IDMEASURE';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EDVES').Index].DataBinding.FieldName := 'EDVES';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EDIZM').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EDIZM').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource10;
          ListFieldNames := 'NAMEEDIZM';
          KeyFieldNames := 'IDEDIZM';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EDIZM').Index].DataBinding.FieldName := 'EDIZM';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('VIDTRANSPORT').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('VIDTRANSPORT').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource11;
          ListFieldNames := 'NAMETRANSPORT';
          KeyFieldNames := 'IDTRANSPORT';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('VIDTRANSPORT').Index].DataBinding.FieldName := 'VIDTRANSPORT';
    end;
  for i:=0 to FrmGrid1.cxGrid1DBTableView1.ColumnCount-1 do
     FrmGrid1.cxGrid1DBTableView1.Columns[i].Width:=80;
  FrmGrid1.Show;
  FrmGrid1.ButtonAdd.DataSource:=nil;
  FrmGrid1.cxGrid1DBTableView1.OptionsData.Editing:=FullDostupBool;
  FrmGrid1.cxGrid1DBTableView1.OptionsData.Deleting:=FullDostupBool;
  FrmGrid1.cxGrid1DBTableView1.OptionsData.Appending:=FullDostupBool;
  //if FullDostupBool=false then FrmGrid1.ButtonDelete.DataSource:=nil;
  FrmGrid1.ButtonDelete.DataSource:=nil;
  FrmGrid1.ToolBarExpertize.Visible:=FullDostupBool;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.ButtonCreateProtokolImportClick(Sender: TObject);
var
  i:integer;
  SQLstr, LockPCName:string;
  LYear, LMonth, LDay: Word;
begin
 Screen.Cursor:=crHourGlass;
 //проверка блокировки номера экспертизы
 ADODataSetLockRegistration.Active:=false;
 SQLstr:='SELECT SETUP.IDVAL FROM SETUP WHERE (((SETUP.IDSET)=1))';
 ADODataSetLockRegistration.CommandText:=SQLstr;
 ADODataSetLockRegistration.Active:=true;
 LockPCName:=ADODataSetLockRegistration.FieldByName('IDVAL').AsString;
 {if LockPCName<>'0' then
    begin
       if LockPCName<>GetComputerNetName then
          begin
             showmessage('В настоящий момент происходит регистрация образцов на другом компьютере: '+LockPCName+','+Chr(13)+'дождитесь завершения ввода данных.'+Chr(13)+'В случае наполадки обратитесь в организационный отдел.');
             ADODataSetLockRegistration.Active:=false;
             Screen.Cursor:=crDefault;
             exit;
          end
       else
          showmessage('Для обеспечения целостности данных не запускайте несколько экземпляров приложения одновременно,'+Chr(13)+'а мастер регистрации образцов закрывайте сразу после ввода данных.');
    end; }
 //блокируем: пишем в базу, что такой-то ПК занял номер
 ADODataSetLockRegistration.Active:=false;
 ADOCommandLockRegistration.CommandText:='UPDATE SETUP SET SETUP.IDVAL = "'+GetComputerNetName+'" WHERE (((SETUP.IDSET)=1))';
 ADOCommandLockRegistration.Execute;
 if FrmDialog1 = NIL then
  begin
    if DateLastDog='' then DateLastDog:=DateToStr(Date-1);
    FrmDialog1 := TFrmDialog1.Create(owner);
    FrmDialog1.DateEdit1.Date:=Date;
    FrmDialog1.DateEdit3.Date:=Date;
    FrmDialog1.DateEdit4.Date:=Date;
  end;
 FrmDialog1.ButtonForward.Caption:='Продолжить >';
 FrmDialog1.ButtonBackward.Enabled:=false;
 FrmDialog1.ButtonCancel.Enabled:=true;
 FrmDialog1.DialogWizard.ActivePage := FrmDialog1.TabSheet1;
 FrmDialog1.LabelPodskazka.Caption:='Шаг 1';
 for i:=0 to FrmDialog1.DialogWizard.PageCount-1 do FrmDialog1.DialogWizard.Pages[i].TabVisible:=false;
 //определяем очередной номер договора и дату последнего
 //DecodeDate(Date-365, LYear, LMonth, LDay);
 DecodeDate(Date, LYear, LMonth, LDay);
 DateTekPerStart:='1/1/'+IntTostr(LYear);
 DecodeDate(Date, LYear, LMonth, LDay);
 DateTekPerFinish:='12/31/'+IntTostr(LYear);
 //showmessage(DateTekPerStart+'*'+DateTekPerFinish);
 //DateTekPerStart:='12/22/2006'; //так было в демо-версии
 //DateTekPerFinish:='12/21/2007'; //так было в демо-версии
 FrmDialog1.ADODataSet12.Active:=false;
 //SQLstr:='SELECT Max(REGISTRATION.NUMPROTOKOL) AS NUM, Max(REGISTRATION.DATEPROTOKOL) AS DATA FROM REGISTRATION WHERE (((REGISTRATION.ISIMPORT)='+IntToStr(FrmDialog1.RadioGroupImport.ItemIndex)+'))';
 //SQLstr:=SQLstr+' HAVING (((Max(REGISTRATION.DATEPROTOKOL))>=#'+DateTekPerStart+'# And (Max(REGISTRATION.DATEPROTOKOL))<#'+DateTekPerFinish+'#)) ORDER BY Max(REGISTRATION.NUMPROTOKOL) DESC;';
 //переделано 29,12,2007
 SQLstr:='SELECT REGISTRATION.NUMPROTOKOL AS NUM, REGISTRATION.DATEPROTOKOL AS DATA FROM REGISTRATION';
 SQLstr:=SQLstr+' WHERE (((REGISTRATION.ISIMPORT)='+IntToStr(FrmDialog1.RadioGroupImport.ItemIndex)+') AND ((REGISTRATION.DATEPROTOKOL)>=#'+DateTekPerStart+'# And (REGISTRATION.DATEPROTOKOL)<#'+DateTekPerFinish+'#)) ORDER BY REGISTRATION.NUMPROTOKOL DESC;';
 FrmDialog1.ADODataSet12.CommandText:=SQLstr;
 FrmDialog1.ADODataSet12.Active:=true;
 if FrmDialog1.ADODataSet12.RecordCount>0 then
  begin
    LastNumProtForAuto:=FrmDialog1.ADODataSet12.FieldByName('NUM').AsInteger+1;
    if FrmDialog1.RadioGroupNumType.ItemIndex=0 then
      FrmDialog1.CalcEdit1.EditValue:=IntToStr(LastNumProtForAuto);
    DateLastDog:=FrmDialog1.ADODataSet12.FieldByName('DATA').AsString;
  end
 else
  begin
    FrmDialog1.CalcEdit1.EditValue:='1';
    DateLastDog:=DateToStr(Date-365);
  end;
 Screen.Cursor:=crDefault;
 FrmDialog1.TabSheet1.Show;
 FrmDialog1.Show;
end;

procedure TFrmMainWindow.ToForm(Sender: TForm);
begin
  Sender.ShowModal;
end;

procedure TFrmMainWindow.ButtonShowClientsClick(Sender: TObject);
begin
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if ADOConnection1.Connected=false then ADOConnection1.Connected:=true;
  ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  ADODataSet1.CommandText:='SELECT * FROM CLIENTS';
  ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCLIENT').Index].Width:=45;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCLIENT').Index].Caption:='Код клиента';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMECLIENT').Index].Width:=270;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMECLIENT').Index].Caption:='Название организации (Ф.И.О. частного лица)';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ABBREVCLIENT').Index].Width:=95;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ABBREVCLIENT').Index].Caption:='Аббревиатура клиента';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ABBREVCLIENT').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      try
        //если номер договора в таблице
        Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DOGNUM').Index].Width:=45;
        Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DOGNUM').Index].Caption:='Номер договора';
      except
      end;
      ADODataSet2.Active:=false;
      ADODataSet2.CommandText:='SELECT * FROM ABBREVIATURA';
      ADODataSet2.Active:=true;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ABBREVCLIENT').Index].Properties) do
        begin
          ListSource := DataSource2;
          ListFieldNames := 'ABBR';
          KeyFieldNames := 'IDABBR';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ABBREVCLIENT').Index].DataBinding.FieldName := 'ABBREVCLIENT';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ADRESSCLIENT').Index].Width:=200;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ADRESSCLIENT').Index].Caption:='Юридический адрес клиента';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Caption:='Код замены';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Visible:=false;
    end;
  UpdStr:='UPDATE REGISTRATION SET ORGANIZATION=';
  UpdKey:='IDCLIENT';
  UpdFld:='ORGANIZATION';
  UpdSpr:='CLIENTS';
  FrmGrid1.Caption:='Клиенты';
  FrmGrid1.Show;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.ExpertMaster(hintmaster: string);
var
  i:integer;
begin
  if FrmDialog2= NIL then FrmDialog2 := TFrmDialog2.Create(owner);
  FrmDialog2.ButtonForward.Caption:='Продолжить >';
  FrmDialog2.ButtonForward.Enabled:=true;
  FrmDialog2.ButtonBackward.Enabled:=false;
  FrmDialog2.ButtonCancel.Enabled:=true;
  FrmDialog2.DialogWizard.ActivePage := FrmDialog2.TabSheet1;
  for i:=0 to FrmDialog2.DialogWizard.PageCount-1 do FrmDialog2.DialogWizard.Pages[i].TabVisible:=false;
  ExpTypeName:=hintmaster;
  FrmDialog2.LabelPodskazka.Caption:='Экспертиза  "'+hintmaster+'"';
  //очистка всех списков объектов
  FrmDialog2.ListViewKarObj.Clear;
  FrmDialog2.ListViewNekarObjOtsRF.Clear;
  FrmDialog2.ListViewNekarObjProchii.Clear;
  //отбор специалистов, которым можно проводить данную экспертизу
  FrmDialog2.ADODataSet1.Active:=false;
  FrmDialog2.ADODataSet1.CommandText:='select * from SPETSIALISTS_LAB WHERE (((SPETSIALISTS_LAB.TYPESPETSIALIST)="'+ExpTypeName+'")) ORDER BY NAMESPETSIALIST';
  FrmDialog2.ADODataSet1.Active:=true;
  FrmDialog2.TabSheet1.Show;
  FrmDialog2.ShowModal;
end;

procedure TFrmMainWindow.ButtonExpVreditelClick(Sender: TObject);
begin
  ExpertMaster(ButtonExpVreditel.Hint);
end;

procedure TFrmMainWindow.ButtonExpGribClick(Sender: TObject);
begin
  ExpertMaster(ButtonExpGrib.Hint);
end;

procedure TFrmMainWindow.ButtonExpBakterClick(Sender: TObject);
begin
  ExpertMaster(ButtonExpBakter.Hint);
end;

procedure TFrmMainWindow.ButtonExpVirusClick(Sender: TObject);
begin
  ExpertMaster(ButtonExpVirus.Hint);
end;

procedure TFrmMainWindow.ButtonExpNematodClick(Sender: TObject);
begin
  ExpertMaster(ButtonExpNematod.Hint);
end;

procedure TFrmMainWindow.ButtonExpSornyClick(Sender: TObject);
begin
  ExpertMaster(ButtonExpSorny.Hint);
end;

procedure TFrmMainWindow.ButtonShowKarObjClick(Sender: TObject);
begin
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if ADOConnection1.Connected=false then ADOConnection1.Connected:=true;
  ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  ADODataSet1.CommandText:='select * from KAROBJECTS;';
  ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;                                
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDOBJ').Index].Width:=45;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDOBJ').Index].Caption:='Код объекта';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('REGOBJ').Index].Width:=120;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('REGOBJ').Index].Caption:='Регистрация в РФ';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Width:=180;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Caption:='Тип объекта';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEOBJ').Index].Width:=350;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEOBJ').Index].Caption:='Наименование объекта, латинское';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEOBJRUS').Index].Width:=350;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEOBJRUS').Index].Caption:='Наименование объекта, русское';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Caption:='Код замены';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].PropertiesClass:=TcxImageComboBoxProperties;
      with FrmGrid1.cxGrid1DBTableView1 do
        begin
          //TcxImageComboBoxProperties(FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Properties).Images:=FrmGrid1.ImageList2;
          TcxImageComboBoxProperties(FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Properties).Items.Add;
          TcxImageComboBoxProperties(FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Properties).Items.Add;
          TcxImageComboBoxProperties(FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Properties).Items.Add;
          TcxImageComboBoxProperties(FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Properties).Items.Add;
          TcxImageComboBoxProperties(FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Properties).Items.Add;
          TcxImageComboBoxProperties(FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Properties).Items.Add;
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Properties).Items[0].Description:='Сорные растения';
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Properties).Items[0].Value:='Сорные растения';
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Properties).Items[1].Description:='Болезни растений, нематодные';
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Properties).Items[1].Value:='Болезни растений, нематодные';
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Properties).Items[2].Description:='Болезни растений, вирусные';
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Properties).Items[2].Value:='Болезни растений, вирусные';
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Properties).Items[3].Description:='Болезни растений, бактериальные';
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Properties).Items[3].Value:='Болезни растений, бактериальные';
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Properties).Items[4].Description:='Болезни растений, грибные';
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Properties).Items[4].Value:='Болезни растений, грибные';
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Properties).Items[5].Description:='Вредители растений';
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEOBJ').Index].Properties).Items[5].Value:='Вредители растений';
        end;
     end;
  UpdStr:='UPDATE EXPERTIZE_DATA SET CODEOBJ=';
  UpdKey:='IDOBJ';
  UpdFld:='CODEOBJ';
  UpdSpr:='KAROBJECTS';
  UpdKarObj:=0;
  FrmGrid1.Show;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.ButtonShowSpetsialistClick(Sender: TObject);
begin
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if ADOConnection1.Connected=false then ADOConnection1.Connected:=true;
  ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  ADODataSet1.CommandText:='select * from SPETSIALISTS;';
  ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDSPETSIALIST').Index].Width:=120;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDSPETSIALIST').Index].Caption:='Код специалиста';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMESPETSIALIST').Index].Width:=220;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMESPETSIALIST').Index].Caption:='Ф.И.О. специалиста';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DOLZHNSPETSIALIST').Index].Width:=200;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DOLZHNSPETSIALIST').Index].Caption:='Должность';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Caption:='Код замены';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Visible:=false;
    end;
  UpdStr:='UPDATE REGISTRATION SET SPETSIALIST=';
  UpdKey:='IDSPETSIALIST';
  UpdFld:='SPETSIALIST';
  UpdSpr:='SPETSIALISTS';
  FrmGrid1.Caption:='Специалисты';
  FrmGrid1.Show;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.ButtonShowUprNadzoraClick(Sender: TObject);
begin
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if FrmMainWindow.ADOConnection1.Connected=false then FrmMainWindow.ADOConnection1.Connected:=true;
  FrmMainWindow.ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  FrmMainWindow.ADODataSet1.CommandText:='select * from UPRSHNADZ;';
  FrmMainWindow.ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=FrmMainWindow.DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDUPRSHNADZ').Index].Width:=27;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDUPRSHNADZ').Index].Caption:='Код управления';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEUPRSHNADZ').Index].Width:=500;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEUPRSHNADZ').Index].Caption:='Полное наименование управления надзора';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMESOKR').Index].Width:=200;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMESOKR').Index].Caption:='Краткое наименование управления надзора';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Caption:='Код замены';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Visible:=false;
    end;
  {UpdStr:='UPDATE REGISTRATION SET UPRSHNADZ=';
  UpdKey:='IDUPRSHNADZ';
  UpdFld:='UPRSHNADZ';
  UpdSpr:='UPRSHNADZ'; }
  UpdStr:='';
  UpdKey:='';
  UpdFld:='';
  UpdSpr:='';
  FrmGrid1.Caption:='Управления';
  FrmGrid1.Show;
  FrmGrid1.cxGrid1DBTableView1.OptionsData.Deleting:=false;
  FrmGrid1.cxGrid1DBTableView1.OptionsData.Editing:=false;
  FrmGrid1.cxGrid1DBTableView1.OptionsData.Appending:=false;
  FrmGrid1.ButtonDelete.Enabled:=false;
  FrmGrid1.ButtonDelete.DataSource:=nil;
  FrmGrid1.ButtonEdit.DataSource:=nil;
  FrmGrid1.ButtonAdd.DataSource:=nil;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.ButtonShowObrazetsClick(Sender: TObject);
begin
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if FrmMainWindow.ADOConnection1.Connected=false then FrmMainWindow.ADOConnection1.Connected:=true;
  FrmMainWindow.ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  FrmMainWindow.ADODataSet1.CommandText:='select * from OBRAZTSI;';
  FrmMainWindow.ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonAdd.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonEdit.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonDelete.DataSource:=FrmMainWindow.DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDOBR').Index].Width:=45;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDOBR').Index].Caption:='Код образца';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEOBR').Index].Width:=500;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEOBR').Index].Caption:='Наименование образца';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Caption:='Код замены';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Visible:=false;
    end;
  UpdStr:='UPDATE REGISTRATION SET NAMEOBRAZTSA=';
  UpdKey:='IDOBR';
  UpdFld:='NAMEOBRAZTSA';
  UpdSpr:='OBRAZTSI';
  FrmGrid1.Caption:='Образцы';
  FrmGrid1.Show;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.ButtonShowPartiiClick(Sender: TObject);
begin
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if FrmMainWindow.ADOConnection1.Connected=false then FrmMainWindow.ADOConnection1.Connected:=true;
  FrmMainWindow.ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  FrmMainWindow.ADODataSet1.CommandText:='select * from PARTII;';
  FrmMainWindow.ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonAdd.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonEdit.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonDelete.DataSource:=FrmMainWindow.DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDPARTII').Index].Width:=45;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDPARTII').Index].Caption:='Код партии';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEPARTII').Index].Width:=150;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEPARTII').Index].Caption:='Название партии';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Caption:='Код замены';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Visible:=false;
    end;
  UpdStr:='UPDATE REGISTRATION SET NAMEPATRY=';
  UpdKey:='IDPARTII';
  UpdFld:='NAMEPATRY';
  UpdSpr:='PARTII';
  FrmGrid1.Caption:='Партии продукции';
  FrmGrid1.Show;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.ButtonShowKlassClick(Sender: TObject);
begin
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if FrmMainWindow.ADOConnection1.Connected=false then FrmMainWindow.ADOConnection1.Connected:=true;
  FrmMainWindow.ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  FrmMainWindow.ADODataSet1.CommandText:='select * from KLASS;';
  FrmMainWindow.ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonAdd.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonEdit.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonDelete.DataSource:=FrmMainWindow.DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDKLASS').Index].Width:=45;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDKLASS').Index].Caption:='Код классификации';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEKLASS').Index].Width:=150;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEKLASS').Index].Caption:='Классификация';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Caption:='Код замены';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Visible:=false;
    end;
  UpdStr:='UPDATE REGISTRATION SET KLASSIFIKATION=';
  UpdKey:='IDKLASS';
  UpdFld:='KLASSIFIKATION';
  UpdSpr:='KLASS';
  FrmGrid1.Caption:='Классификации';
  FrmGrid1.Show;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.ButtonShowProishozhdClick(Sender: TObject);
begin
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if FrmMainWindow.ADOConnection1.Connected=false then FrmMainWindow.ADOConnection1.Connected:=true;
  FrmMainWindow.ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  FrmMainWindow.ADODataSet1.CommandText:='select * from PROISHOZHD;';
  FrmMainWindow.ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonAdd.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonEdit.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonDelete.DataSource:=FrmMainWindow.DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDPROISH').Index].Width:=45;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDPROISH').Index].Caption:='Код';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEPROISH').Index].Width:=150;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEPROISH').Index].Caption:='Происхождение образца';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Caption:='Код замены';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Visible:=false;
    end;
  UpdStr:='UPDATE REGISTRATION SET PROISHOZHD=';
  UpdKey:='IDPROISH';
  UpdFld:='PROISHOZHD';
  UpdSpr:='PROISHOZHD';
  FrmGrid1.Caption:='Происхождение';
  FrmGrid1.Show;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.ButtonShowEdIzmClick(Sender: TObject);
begin
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if FrmMainWindow.ADOConnection1.Connected=false then FrmMainWindow.ADOConnection1.Connected:=true;
  FrmMainWindow.ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  FrmMainWindow.ADODataSet1.CommandText:='select * from EDIZM;';
  FrmMainWindow.ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonAdd.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonEdit.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonDelete.DataSource:=FrmMainWindow.DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDEDIZM').Index].Width:=45;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDEDIZM').Index].Caption:='Код единицы';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEEDIZM').Index].Width:=150;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEEDIZM').Index].Caption:='Название';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Caption:='Код замены';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Visible:=false;
    end;
  UpdStr:='';
  UpdFld:='';
  UpdSpr:='';
  FrmGrid1.Caption:='Единицы измерения';
  FrmGrid1.Show;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.ButtonShowVesClick(Sender: TObject);
begin
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if FrmMainWindow.ADOConnection1.Connected=false then FrmMainWindow.ADOConnection1.Connected:=true;
  FrmMainWindow.ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  FrmMainWindow.ADODataSet1.CommandText:='select * from VES;';
  FrmMainWindow.ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=FrmMainWindow.DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDMEASURE').Index].Width:=45;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDMEASURE').Index].Caption:='Код единицы веса';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEMEASURE').Index].Width:=150;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEMEASURE').Index].Caption:='Единица веса';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Caption:='Код замены';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Visible:=false;
    end;
  UpdStr:='';
  UpdFld:='';
  UpdSpr:='';
  FrmGrid1.Caption:='Единицы веса';
  FrmGrid1.Show;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.ButtonShowTransportClick(Sender: TObject);
begin
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if FrmMainWindow.ADOConnection1.Connected=false then FrmMainWindow.ADOConnection1.Connected:=true;
  FrmMainWindow.ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  FrmMainWindow.ADODataSet1.CommandText:='select * from TRANSPORT;';
  FrmMainWindow.ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=FrmMainWindow.DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDTRANSPORT').Index].Width:=45;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDTRANSPORT').Index].Caption:='Код';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMETRANSPORT').Index].Width:=150;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMETRANSPORT').Index].Caption:='Вид транспортировки';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Caption:='Код замены';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Visible:=false;
    end;
  UpdStr:='UPDATE REGISTRATION SET VIDTRANSPORT=';
  UpdKey:='IDTRANSPORT';
  UpdFld:='VIDTRANSPORT';
  UpdSpr:='TRANSPORT';
  FrmGrid1.Caption:='Транспорт';
  FrmGrid1.Show;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.ButtonNekarObjProcheeClick(Sender: TObject);
begin
  PopupMenu2.Popup(Left+AdvToolBar10.Left+13,Top+138);
end;

procedure TFrmMainWindow.ButtonComputersClick(Sender: TObject);
begin
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if FrmMainWindow.ADOConnection1.Connected=false then FrmMainWindow.ADOConnection1.Connected:=true;
  FrmMainWindow.ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  FrmMainWindow.ADODataSet1.CommandText:='select * from USERS;';
  FrmMainWindow.ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonAdd.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonEdit.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonDelete.DataSource:=FrmMainWindow.DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCOMP').Index].Width:=45;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCOMP').Index].Caption:='Номер компьютера';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMECOMP').Index].Width:=150;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMECOMP').Index].Caption:='Название компьютера';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEUSER').Index].Width:=150;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEUSER').Index].Caption:='Имя пользователя';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('FULLDOSTUP').Index].Width:=30;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('FULLDOSTUP').Index].Caption:='Администрирование';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EXPVREDITEL').Index].Width:=30;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EXPVREDITEL').Index].Caption:='Вредители, доступ к экспертизе';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EXPGRIB').Index].Width:=30;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EXPGRIB').Index].Caption:='Грибы, доступ к экспертизе';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EXPBAKTER').Index].Width:=30;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EXPBAKTER').Index].Caption:='Бактерии, доступ к экспертизе';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EXPVIRUS').Index].Width:=30;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EXPVIRUS').Index].Caption:='Вирусы, доступ к экспертизе';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EXPNEMATOD').Index].Width:=30;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EXPNEMATOD').Index].Caption:='Нематоды, доступ к экспертизе';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EXPSORN').Index].Width:=30;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('EXPSORN').Index].Caption:='Сорняки, доступ к экспертизе';
    end;
  FrmGrid1.Show;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.ButtonProtokolTemplateClick(Sender: TObject);
begin
  ButtonShowExcel.Enabled:=true;
  try
    ExcelApp := GetActiveOleObject('Excel.Application');   //проверяем, нет ли запущенного Excel
  except
    on EOLESysError do ExcelApp:= CreateOleObject('Excel.Application');   //если нет, то запускаем
  end;
  ExcelApp.Application.EnableEvents:= false; // так будет быстрее
  ExcelApp.Application.DisplayAlerts:= true; // спрашивать в Excel сохранять ли файл после изменений
  ExcelApp.Visible:=true;
  Workbook := ExcelApp.WorkBooks.Open(ExtractFilePath(Application.ExeName)+'Templates\Протокол.xls');
  ExcelApp.DisplayFullScreen:=true;
  ExcelApp.DisplayFullScreen:=false;
end;

procedure TFrmMainWindow.ButtonSvvoTemplateClick(Sender: TObject);
begin
  ButtonShowExcel.Enabled:=true;
  try
    ExcelApp := GetActiveOleObject('Excel.Application');   //проверяем, нет ли запущенного Excel
  except
    on EOLESysError do ExcelApp:= CreateOleObject('Excel.Application');   //если нет, то запускаем
  end;
  ExcelApp.Application.EnableEvents:= false; // так будет быстрее
  ExcelApp.Application.DisplayAlerts:= true; // спрашивать в Excel сохранять ли файл после изменений
  ExcelApp.Visible:=true;
  Workbook := ExcelApp.WorkBooks.Open(ExtractFilePath(Application.ExeName)+'Templates\Свидетельство.xls');
  ExcelApp.DisplayFullScreen:=true;
  ExcelApp.DisplayFullScreen:=false;
end;

procedure TFrmMainWindow.ValidateAccessDB;
var 
  lDBpathName : String;
  TUsers: TADOTable;
begin 
  if FileExists(ExtractFileDir(Application.ExeName) + '\Karex.mdb' ) then
    lDBPathName := ExtractFileDir(Application.ExeName) + '\Karex.mdb';
  ADOConnection1.ConnectionString :=
      'Provider=Microsoft.Jet.OLEDB.4.0;' + 
      'Data Source=' + lDBPathName + ';' + 
      'Persist Security Info=False;' + 
      'Jet OLEDB:Database Password=epson';
   ADOConnection1.Connected:=true;
end;

procedure TFrmMainWindow.FormCreate(Sender: TObject);
var
  LoginName: string;
  REG: TRegistry;
begin
  Application.CreateForm(TFrmLogo, FrmLogo);
  FrmLogo.ShowModal;
  validateAccessDB;
  Timer1.Enabled:=false;
  try
    LoginName:=GetLoginName;
    dxStatusBar1.Panels[1].Text:=LoginName;
    dxStatusBar1.Panels[1].Width:=Length(LoginName)*7;
    ADODataSet1.Active:=false;
    ADODataSet1.CommandText:='SELECT * FROM USERS WHERE (((NAMECOMP)="'+GetComputerNetName+'"));';
    ADODataSet1.Active:=true;
    if ADODataSet1.RecordCount>0 then
      begin
        dxStatusBar1.Panels[0].Text:=GetComputerNetName;
        ButtonExpVreditel.Enabled:=StrToBool(ADODataSet1.FieldByName('EXPVREDITEL').AsString);
        ButtonExpGrib.Enabled:=StrToBool(ADODataSet1.FieldByName('EXPGRIB').AsString);
        ButtonExpBakter.Enabled:=StrToBool(ADODataSet1.FieldByName('EXPBAKTER').AsString);
        ButtonExpVirus.Enabled:=StrToBool(ADODataSet1.FieldByName('EXPVIRUS').AsString);
        ButtonExpNematod.Enabled:=StrToBool(ADODataSet1.FieldByName('EXPNEMATOD').AsString);
        ButtonExpSorny.Enabled:=StrToBool(ADODataSet1.FieldByName('EXPSORN').AsString);
        ButtonCreateProtokolImport.Enabled:=StrToBool(ADODataSet1.FieldByName('FULLDOSTUP').AsString);
        ButtonComputers.Enabled:=StrToBool(ADODataSet1.FieldByName('FULLDOSTUP').AsString);
        ButtonProtokolTemplate.Enabled:=StrToBool(ADODataSet1.FieldByName('FULLDOSTUP').AsString);
        ButtonSvvoTemplate.Enabled:=StrToBool(ADODataSet1.FieldByName('FULLDOSTUP').AsString);
        ButtonSvodTemplate.Enabled:=StrToBool(ADODataSet1.FieldByName('FULLDOSTUP').AsString);
        ButtonSvodOtch.Enabled:=StrToBool(ADODataSet1.FieldByName('FULLDOSTUP').AsString);
        FullDostupBool:=StrToBool(ADODataSet1.FieldByName('FULLDOSTUP').AsString);
        ButtonSvvo2Template.Enabled:=FullDostupBool;
        ButtonSvvo3Template.Enabled:=FullDostupBool;
        ButtonAktTemplate.Enabled:=FullDostupBool;
        ButtonUnblockRegProtokol.Enabled:=FullDostupBool;
        //доступ к контекстным меню
        N1.Enabled:=ButtonExpVreditel.Enabled;
        N2.Enabled:=ButtonExpGrib.Enabled;
        N3.Enabled:=ButtonExpBakter.Enabled;
        N4.Enabled:=ButtonExpVirus.Enabled;
        N5.Enabled:=ButtonExpNematod.Enabled;
        N6.Enabled:=ButtonExpSorny.Enabled;
        N21.Enabled:=ButtonExpVreditel.Enabled;
        N22.Enabled:=ButtonExpGrib.Enabled;
        N23.Enabled:=ButtonExpBakter.Enabled;
        N24.Enabled:=ButtonExpVirus.Enabled;
        N25.Enabled:=ButtonExpNematod.Enabled;
        N26.Enabled:=ButtonExpSorny.Enabled;
      end
    else
      begin
        dxStatusBar1.Panels[0].Text:='Пользователь не зарегистрирован';
        AdvToolBar1.Enabled:=false;
        AdvToolBar13.Enabled:=false;
        AdvToolBar4.Enabled:=false;
        AdvPage3.TabVisible:=false;
        FullDostupBool:=false;
        ButtonExpVreditel.Enabled:=false;
        ButtonExpGrib.Enabled:=false;
        ButtonExpBakter.Enabled:=false;
        ButtonExpVirus.Enabled:=false;
        ButtonExpNematod.Enabled:=false;
        ButtonExpSorny.Enabled:=false;
        ButtonCreateProtokolImport.Enabled:=false;
        ButtonComputers.Enabled:=false;
        ButtonProtokolTemplate.Enabled:=false;
        ButtonSvvoTemplate.Enabled:=false;
        ButtonNewReg.Enabled:=true;
      end;
  except
  end;
  dxStatusBar1.Panels[2].Text:='  Сегодня '+DateToStr(Date);
  fc:=0;
  //открытие Excel после занесения протокола
  REG := TRegistry.Create;
  REG.RootKey:=HKEY_CURRENT_USER;
  if REG.OpenKey('Software\Karex\Local\Settings',true) then;
  if REG.ValueExists('OpenExcelAfterProtEnt') then Oeape:=REG.ReadBool('OpenExcelAfterProtEnt') else Oeape:=true;
  if REG.ValueExists('MultiUserReg') then MUReg:=REG.ReadBool('MultiUserReg') else MUReg:=false;
  if REG.ValueExists('ViewStatistics') then ViewStatistic:=REG.ReadBool('ViewStatistics') else ViewStatistic:=false;
  REG.CloseKey;
  REG.Free;
  FrmLogo.Free;
end;

procedure TFrmMainWindow.ButtonShowSpetsLabClick(Sender: TObject);
var
  i:integer;
begin
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if ADOConnection1.Connected=false then ADOConnection1.Connected:=true;
  ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  ADODataSet1.CommandText:='select * from SPETSIALISTS_LAB;';
  ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDSPETSIALIST').Index].Width:=120;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDSPETSIALIST').Index].Caption:='Код специалиста';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMESPETSIALIST').Index].Width:=200;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMESPETSIALIST').Index].Caption:='Ф.И.О. специалиста';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DOLZHNSPETSIALIST').Index].Width:=150;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DOLZHNSPETSIALIST').Index].Caption:='Должность';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPESPETSIALIST').Index].Width:=200;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPESPETSIALIST').Index].Caption:='Зона экспертиз';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Caption:='Код замены';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPESPETSIALIST').Index].PropertiesClass:=TcxImageComboBoxProperties;
    end;
  for i:=0 to 6 do TcxImageComboBoxProperties(FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPESPETSIALIST').Index].Properties).Items.Add;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPESPETSIALIST').Index].Properties).Items[0].Description:='Вредители растений';
      TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPESPETSIALIST').Index].Properties).Items[0].Value:='Вредители растений';
      TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPESPETSIALIST').Index].Properties).Items[1].Description:='Болезни растений, грибные';
      TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPESPETSIALIST').Index].Properties).Items[1].Value:='Болезни растений, грибные';
      TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPESPETSIALIST').Index].Properties).Items[2].Description:='Болезни растений, бактериальные';
      TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPESPETSIALIST').Index].Properties).Items[2].Value:='Болезни растений, бактериальные';
      TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPESPETSIALIST').Index].Properties).Items[3].Description:='Болезни растений, вирусные';
      TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPESPETSIALIST').Index].Properties).Items[3].Value:='Болезни растений, вирусные';
      TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPESPETSIALIST').Index].Properties).Items[4].Description:='Болезни растений, нематодные';
      TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPESPETSIALIST').Index].Properties).Items[4].Value:='Болезни растений, нематодные';
      TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPESPETSIALIST').Index].Properties).Items[5].Description:='Сорные растения';
      TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPESPETSIALIST').Index].Properties).Items[5].Value:='Сорные растения';
      TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPESPETSIALIST').Index].Properties).Items[6].Description:='Отбор образцов';
      TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPESPETSIALIST').Index].Properties).Items[6].Value:='Отбор образцов';
    end;
  UpdStr:='';
  UpdFld:='';
  UpdSpr:='';
  FrmGrid1.Caption:='Специалисты';
  FrmGrid1.Show;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.ButtonShowAbbreviaturaClick(Sender: TObject);
begin
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if ADOConnection1.Connected=false then ADOConnection1.Connected:=true;
  ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  ADODataSet1.CommandText:='select * from ABBREVIATURA;';
  ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDABBR').Index].Width:=110;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDABBR').Index].Caption:='Код аббревиатуры';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ABBR').Index].Width:=270;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ABBR').Index].Caption:='Аббревиатура';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Caption:='Код замены';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Visible:=false;
    end;
  UpdStr:='';
  FrmGrid1.Caption:='Аббревиатуры';
  FrmGrid1.Show;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.ButtonShowPunktsClick(Sender: TObject);
begin
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if FrmMainWindow.ADOConnection1.Connected=false then FrmMainWindow.ADOConnection1.Connected:=true;
  FrmMainWindow.ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  FrmMainWindow.ADODataSet1.CommandText:='select * from KARPUNKTS;';
  FrmMainWindow.ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonAdd.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonEdit.DataSource:=FrmMainWindow.DataSource1;
  FrmGrid1.ButtonDelete.DataSource:=FrmMainWindow.DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDKARPUNKT').Index].Width:=45;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDKARPUNKT').Index].Caption:='Код пункта';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEKARPUNKT').Index].Width:=150;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMEKARPUNKT').Index].Caption:='Наименование пункта карантина';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Caption:='Код замены';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Visible:=false;
    end;
  UpdStr:='UPDATE REGISTRATION SET PUNKTKARRAST=';
  UpdKey:='IDKARPUNKT';
  UpdFld:='PUNKTKARRAST';
  UpdSpr:='KARPUNKTS';
  FrmGrid1.Caption:='Пункты карантина';
  FrmGrid1.Show;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.ButtonShowDocVidsClick(Sender: TObject);
begin
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if FrmMainWindow.ADOConnection1.Connected=false then FrmMainWindow.ADOConnection1.Connected:=true;
  FrmMainWindow.ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  FrmMainWindow.ADODataSet1.CommandText:='select * from DOCVID;';
  FrmMainWindow.ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=FrmMainWindow.DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDDOCVIDS').Index].Width:=45;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDDOCVIDS').Index].Caption:='Код вида';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DOCVIDS').Index].Width:=600;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DOCVIDS').Index].Caption:='Полное наименование вида документа';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DOCSOKR').Index].Width:=60;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DOCSOKR').Index].Caption:='Сокращённое наименование вида документа';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Caption:='Код замены';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Visible:=false;
    end;
  UpdStr:='';
  UpdKey:='';
  UpdFld:='';
  UpdSpr:='';
  FrmGrid1.Caption:='Виды документов';
  FrmGrid1.Show;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.NekarObjOtsRF(typeobj: string);
begin
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if ADOConnection1.Connected=false then ADOConnection1.Connected:=true;
  ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  ADODataSet1.CommandText:='select * from NEKAROBJECTS_NEZARRF WHERE (((NEKAROBJECTS_NEZARRF.TYPENEKAROBJ)="'+typeobj+'"));';
  ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDNEKAROBJ').Index].Width:=45;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDNEKAROBJ').Index].Caption:='Код объекта';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPENEKAROBJ').Index].Width:=120;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPENEKAROBJ').Index].Caption:='Тип объекта';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPENEKAROBJ').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPENEKAROBJ').Index].Hidden:=true;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADLAT').Index].Width:=350;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADLAT').Index].Caption:='Отряд, латинское название';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADRUS').Index].Width:=350;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADRUS').Index].Caption:='Отряд, русское название';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('FAMILYLAT').Index].Width:=350;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('FAMILYLAT').Index].Caption:='Семейство, латинское название';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('FAMILYRUS').Index].Width:=350;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('FAMILYRUS').Index].Caption:='Семейство, русское название';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMELAT').Index].Width:=350;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMELAT').Index].Caption:='Латинское название';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMERUS').Index].Width:=350;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMERUS').Index].Caption:='Русское название';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Caption:='Код замены';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Visible:=false;
      if typeobj='Сорные растения' then
        begin
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADLAT').Index].Visible:=false;
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADRUS').Index].Visible:=false;
        end;
      if typeobj='Болезни растений, грибные' then
        begin
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADLAT').Index].Caption:='Класс, латинское название';
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADRUS').Index].Visible:=false;
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('FAMILYRUS').Index].Visible:=false;
        end;
      if typeobj='Болезни растений, бактериальные' then
        begin
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADLAT').Index].Visible:=false;
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADRUS').Index].Visible:=false;
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('FAMILYLAT').Index].Visible:=false;
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('FAMILYRUS').Index].Visible:=false;
        end;
      if typeobj='Болезни растений, нематодные' then
        begin
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADLAT').Index].Caption:='Род, латинское название';
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADRUS').Index].Visible:=false;
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('FAMILYRUS').Index].Caption:='Подсемейство, латинское название';
        end;
    end;
  //значение по умолчанию
  DefaultVal:=typeobj;
  {
  UpdStr - начало строки замены в таблице данных
  UpdFld - поле в таблице данных, которое меняется
  UpdSpr - таблица справочника
  UpdKey - поле ключ в таблице справочника
  UpdKarObj - тип объекта
  }
  UpdStr:='UPDATE EXPERTIZE_DATA SET CODEOBJ=';
  UpdFld:='CODEOBJ';
  UpdSpr:='NEKAROBJECTS_NEZARRF';
  UpdKey:='IDNEKAROBJ';
  UpdKarObj:=1;
  FrmGrid1.Show;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.NekarObjProchie(typeobj: string);
begin
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if ADOConnection1.Connected=false then ADOConnection1.Connected:=true;
  ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  ADODataSet1.CommandText:='select * from NEKAROBJECTS_PROCHIE WHERE (((NEKAROBJECTS_PROCHIE.TYPENEKAROBJ)="'+typeobj+'"));';
  ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDNEKAROBJ').Index].Width:=45;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDNEKAROBJ').Index].Caption:='Код объекта';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPENEKAROBJ').Index].Width:=120;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPENEKAROBJ').Index].Caption:='Тип объекта';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPENEKAROBJ').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPENEKAROBJ').Index].Hidden:=true;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADLAT').Index].Width:=350;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADLAT').Index].Caption:='Отряд, латинское название';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADRUS').Index].Width:=350;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADRUS').Index].Caption:='Отряд, русское название';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('FAMILYLAT').Index].Width:=350;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('FAMILYLAT').Index].Caption:='Семейство, латинское название';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('FAMILYRUS').Index].Width:=350;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('FAMILYRUS').Index].Caption:='Семейство, русское название';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMELAT').Index].Width:=350;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMELAT').Index].Caption:='Латинское название';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMERUS').Index].Width:=350;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMERUS').Index].Caption:='Русское название';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Caption:='Код замены';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDCHANGE').Index].Visible:=false;
      if typeobj='Сорные растения' then
        begin
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADLAT').Index].Visible:=false;
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADRUS').Index].Visible:=false;
        end;
      if typeobj='Болезни растений, грибные' then
        begin
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADLAT').Index].Caption:='Класс, латинское название';
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADRUS').Index].Visible:=false;
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('FAMILYRUS').Index].Visible:=false;
        end;
      if typeobj='Болезни растений, бактериальные' then
        begin
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADLAT').Index].Visible:=false;
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADRUS').Index].Visible:=false;
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('FAMILYLAT').Index].Visible:=false;
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('FAMILYRUS').Index].Visible:=false;
        end;
      if typeobj='Болезни растений, нематодные' then
        begin
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADLAT').Index].Caption:='Род, латинское название';
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('OTRYADRUS').Index].Visible:=false;
           Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('FAMILYRUS').Index].Caption:='Подсемейство, латинское название';
        end;
    end;
  //значение по умолчанию
  DefaultVal:=typeobj;
  {
  UpdStr - начало строки замены в таблице данных
  UpdFld - поле в таблице данных, которое меняется
  UpdSpr - таблица справочника
  UpdKey - поле ключ в таблице справочника
  UpdKarObj - тип объекта
  }
  UpdStr:='UPDATE EXPERTIZE_DATA SET CODEOBJ=';
  UpdFld:='CODEOBJ';
  UpdSpr:='NEKAROBJECTS_PROCHIE';
  UpdKey:='IDNEKAROBJ';
  UpdKarObj:=2;
  FrmGrid1.Show;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.N1Click(Sender: TObject);
begin
  NekarObjOtsRF(N1.Caption);
end;

procedure TFrmMainWindow.N2Click(Sender: TObject);
begin
  NekarObjOtsRF(N2.Caption);
end;

procedure TFrmMainWindow.N3Click(Sender: TObject);
begin
  NekarObjOtsRF(N3.Caption);
end;

procedure TFrmMainWindow.N4Click(Sender: TObject);
begin
  NekarObjOtsRF(N4.Caption);
end;

procedure TFrmMainWindow.N5Click(Sender: TObject);
begin
  NekarObjOtsRF(N5.Caption);
end;

procedure TFrmMainWindow.N6Click(Sender: TObject);
begin
  NekarObjOtsRF(N6.Caption);
end;

procedure TFrmMainWindow.ADODataSet1NewRecord(DataSet: TDataSet);
begin
  if DefaultVal<>'' then FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPENEKAROBJ').Index].EditValue:=DefaultVal;
end;

procedure TFrmMainWindow.SelSpetsialistForExp(Sender: TObject);
begin
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if ADOConnection1.Connected=false then ADOConnection1.Connected:=true;
  ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  ADODataSet1.CommandText:='select * from SPETSIALISTS_LAB WHERE (((SPETSIALISTS_LAB.TYPESPETSIALIST)="'+ExpTypeName+'")) ORDER BY NAMESPETSIALIST';
  ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDSPETSIALIST').Index].Width:=120;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDSPETSIALIST').Index].Caption:='Код специалиста';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMESPETSIALIST').Index].Width:=200;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMESPETSIALIST').Index].Caption:='Ф.И.О. специалиста';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DOLZHNSPETSIALIST').Index].Width:=150;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DOLZHNSPETSIALIST').Index].Caption:='Должность';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPESPETSIALIST').Index].Width:=200;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPESPETSIALIST').Index].Caption:='Зона экспертиз';
    end;
  FrmGrid1.Show;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.ButtonExpFindClick(Sender: TObject);
var
  SQLStr, ObjTbl, IDField: string;
  i: integer;
begin
 FrmPeriodExp.RadioGroupParamVibor.ItemIndex:=0;
 FrmPeriodExp.Showmodal;
 if PerStart='' then exit;
 if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
 if FrmDialog1= NIL then FrmDialog1 := TFrmDialog1.Create(owner);
 if ADOConnection1.Connected=false then ADOConnection1.Connected:=true;
 ADODataSet1.Active:=false;
 FrmGrid1.cxGrid1DBTableView1.ClearItems;
 if FrmPeriodExp.RadioGroupParamVibor.ItemIndex=1 then
  begin //если подробный отчёт, но по конкретному типу -------------------------------------------------
  case FrmPeriodExp.RadioGroupObj.ItemIndex of
    0: begin ObjTbl:='KAROBJECTS'; IDField:='IDOBJ'; end;
    1: begin ObjTbl:='NEKAROBJECTS_NEZARRF'; IDField:='IDNEKAROBJ'; end;
    2: begin ObjTbl:='NEKAROBJECTS_PROCHIE'; IDField:='IDNEKAROBJ'; end;
  end;
  SQLStr:='SELECT EXPERTIZE_REESTR.IDEXPERTIZE, EXPERTIZE_REESTR.NUMPROTOKOL, EXPERTIZE_REESTR.SPETSIALIST, EXPERTIZE_DATA.TYPEEXPERTIZE, EXPERTIZE_DATA.CODEOBJ,';
  SQLStr:=SQLStr+' EXPERTIZE_DATA.SOSTOYANIE, EXPERTIZE_DATA.PRIMECH, EXPERTIZE_REESTR.DATEEXPERTIZE, REGISTRATION.NUMPROTOKOL AS NP, REGISTRATION.DATEPROTOKOL, REGISTRATION.ORGANIZATION,';
  SQLStr:=SQLStr+' EXPERTIZE_DATA.IDROW, '+ObjTbl+'.'+IDField+', REGISTRATION.IDREGISTRATION';
  SQLStr:=SQLStr+' FROM ((EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE) LEFT JOIN '+ObjTbl+' ON EXPERTIZE_DATA.CODEOBJ = '+ObjTbl+'.'+IDField+') LEFT JOIN REGISTRATION ON';
  SQLStr:=SQLStr+' EXPERTIZE_REESTR.NUMPROTOKOL = REGISTRATION.IDREGISTRATION';
  SQLStr:=SQLStr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)='+IntToStr(FrmPeriodExp.RadioGroupObj.ItemIndex)+') AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)='+IntToStr(FrmPeriodExp.RadioGroupImport.ItemIndex)+'));';
  ADODataSet1.CommandText:=SQLStr;
  ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=DataSource1;
  FrmGrid1.cxGrid1DBTableView1.DataController.DetailKeyFieldNames:='IDEXPERTIZE';
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  for i:=0 to FrmGrid1.cxGrid1DBTableView1.ColumnCount-1 do
     FrmGrid1.cxGrid1DBTableView1.Columns[i].Width:=80;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDEXPERTIZE').Index].Caption:='Код экспертизы';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDEXPERTIZE').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NP').Index].Caption:='Номер протокола';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NP').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].Caption:='Индекс протокола';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('SPETSIALIST').Index].Caption:='Ф.И.О. специалиста';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEEXPERTIZE').Index].Caption:='Тип экспертизы';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEEXPERTIZE').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('CODEOBJ').Index].Caption:='Наименование объекта';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('CODEOBJ').Index].Width:=200;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('SOSTOYANIE').Index].Caption:='Состояние объекта';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PRIMECH').Index].Caption:='Примечание объекта';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DATEEXPERTIZE').Index].Caption:='Дата экспертизы';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DATEPROTOKOL').Index].Caption:='Дата протокола';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DATEPROTOKOL').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].Caption:='Клиент';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDROW').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDROW').Index].Hidden:=true;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName(IDField).Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName(IDField).Index].Hidden:=true;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].Hidden:=true;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('SPETSIALIST').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('SPETSIALIST').Index].Properties) do
        begin
          ListSource := FrmDialog2.DataSource1;
          ListFieldNames := 'NAMESPETSIALIST';
          KeyFieldNames := 'IDSPETSIALIST';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      ADODataSet2.Active:=false;
      ADODataSet2.CommandText:='select * from '+ObjTbl;
      ADODataSet2.Active:=true;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('CODEOBJ').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('CODEOBJ').Index].Properties) do
        begin
          ListSource := DataSource2;
          case FrmPeriodExp.RadioGroupObj.ItemIndex of
            0: begin
                ListFieldNames := 'NAMEOBJ;NAMEOBJRUS';
                KeyFieldNames := 'IDOBJ';
               end;
            1,2: begin
                ListFieldNames := 'NAMELAT;NAMERUS';
                KeyFieldNames := 'IDNEKAROBJ';
               end;
          end;
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource4;
          ListFieldNames := 'NAMECLIENT';
          KeyFieldNames := 'IDCLIENT';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
    end;
    FrmGrid1.Show;
    FrmGrid1.ButtonAdd.DataSource:=nil;
    FrmGrid1.cxGrid1DBTableView1.OptionsData.Editing:=FullDostupBool;
    FrmGrid1.cxGrid1DBTableView1.OptionsData.Deleting:=FullDostupBool;
    FrmGrid1.cxGrid1DBTableView1.OptionsData.Appending:=FullDostupBool;
    //if FullDostupBool=false then FrmGrid1.ButtonDelete.DataSource:=nil;
    FrmGrid1.ButtonDelete.DataSource:=nil;
  end // если подробный отчёт, но по конкретному типу ----------------------------------------------
  else if FrmPeriodExp.RadioGroupParamVibor.ItemIndex=0 then
  begin // если сводный отчёт по всем типам
      SQLStr:='SELECT EXPERTIZE_REESTR.IDEXPERTIZE, EXPERTIZE_REESTR.ISIMPORT, EXPERTIZE_REESTR.DATEEXPERTIZE, REGISTRATION.NUMPROTOKOL, SPETSIALISTS_LAB.NAMESPETSIALIST, SPETSIALISTS_LAB.DOLZHNSPETSIALIST';
      SQLStr:=SQLStr+' FROM (EXPERTIZE_REESTR LEFT JOIN SPETSIALISTS_LAB ON EXPERTIZE_REESTR.SPETSIALIST = SPETSIALISTS_LAB.IDSPETSIALIST) LEFT JOIN REGISTRATION ON EXPERTIZE_REESTR.NUMPROTOKOL = REGISTRATION.IDREGISTRATION';
      SQLStr:=SQLStr+' WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'#-1 And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#+1))';
      SQLStr:=SQLStr+' ORDER BY REGISTRATION.NUMPROTOKOL DESC';
      ADODataSet1.CommandText:=SQLStr;
      ADODataSet1.Active:=true;
      FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=DataSource1;
      FrmGrid1.cxGrid1DBTableView1.DataController.DetailKeyFieldNames:='IDEXPERTIZE';
      with FrmGrid1.cxGrid1DBTableView1.DataController do
        CreateAllItems;
      for i:=0 to FrmGrid1.cxGrid1DBTableView1.ColumnCount-1 do
      FrmGrid1.cxGrid1DBTableView1.Columns[i].Width:=130;
      with FrmGrid1.cxGrid1DBTableView1 do
        begin
          Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDEXPERTIZE').Index].Caption:='Код экспертизы';
          Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Caption:='Происхождение';
          Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DATEEXPERTIZE').Index].Caption:='Дата экспертизы';
          Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].Caption:='Номер протокола';
          Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NAMESPETSIALIST').Index].Caption:='Ф.И.О. специалиста';
          Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DOLZHNSPETSIALIST').Index].Caption:='Должность специалиста';
          Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].PropertiesClass:=TcxImageComboBoxProperties;
          with FrmGrid1.cxGrid1DBTableView1 do
            begin
              TcxImageComboBoxProperties(FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Images:=FrmGrid1.ImageList2;
              TcxImageComboBoxProperties(FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items.Add;
              TcxImageComboBoxProperties(FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items.Add;
              TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items[0].Description:='Импортный';
              TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items[0].Value:=0;
              TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items[1].Description:='Отечественный';
              TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items[1].Value:=1;
            end;
        end;
      FrmGrid1.Show;
      FrmGrid1.ButtonAdd.DataSource:=nil;
      FrmGrid1.cxGrid1DBTableView1.OptionsData.Editing:=false;
      FrmGrid1.cxGrid1DBTableView1.OptionsData.Deleting:=false;
      FrmGrid1.cxGrid1DBTableView1.OptionsData.Appending:=false;
      FrmGrid1.ButtonDelete.DataSource:=nil;
  end
  else
  //если идёт выборка по коду
  begin // -------------------------------------------------
  case FrmPeriodExp.RadioGroupObj.ItemIndex of
    0: begin ObjTbl:='KAROBJECTS'; IDField:='IDOBJ'; end;
    1: begin ObjTbl:='NEKAROBJECTS_NEZARRF'; IDField:='IDNEKAROBJ'; end;
    2: begin ObjTbl:='NEKAROBJECTS_PROCHIE'; IDField:='IDNEKAROBJ'; end;
  end;
  SQLStr:='SELECT EXPERTIZE_REESTR.IDEXPERTIZE, EXPERTIZE_REESTR.NUMPROTOKOL, EXPERTIZE_REESTR.SPETSIALIST, EXPERTIZE_DATA.TYPEEXPERTIZE, EXPERTIZE_DATA.CODEOBJ,';
  SQLStr:=SQLStr+' EXPERTIZE_DATA.SOSTOYANIE, EXPERTIZE_DATA.PRIMECH, EXPERTIZE_REESTR.DATEEXPERTIZE, REGISTRATION.NUMPROTOKOL AS NP, REGISTRATION.DATEPROTOKOL, REGISTRATION.ORGANIZATION,';
  SQLStr:=SQLStr+' EXPERTIZE_DATA.IDROW, '+ObjTbl+'.'+IDField+', REGISTRATION.IDREGISTRATION';
  SQLStr:=SQLStr+' FROM ((EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE) LEFT JOIN '+ObjTbl+' ON EXPERTIZE_DATA.CODEOBJ = '+ObjTbl+'.'+IDField+') LEFT JOIN REGISTRATION ON';
  SQLStr:=SQLStr+' EXPERTIZE_REESTR.NUMPROTOKOL = REGISTRATION.IDREGISTRATION';
  SQLStr:=SQLStr+' WHERE ((REGISTRATION.NUMPROTOKOL)='+IntToStr(FrmPeriodExp.CalcEditCode.EditValue)+' and ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'#-1 And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#+1))';   // where
  ADODataSet1.CommandText:=SQLStr;
  ADODataSet1.Active:=true;
  if ADODataSet1.RecordCount=0 then
     begin
        showmessage('По данному протоколу нет записей об обнаруженных объектах!');
        exit;
     end;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=DataSource1;
  FrmGrid1.cxGrid1DBTableView1.DataController.DetailKeyFieldNames:='IDEXPERTIZE';
  FrmGrid1.ButtonDeleteProtokol.Visible:=true;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  for i:=0 to FrmGrid1.cxGrid1DBTableView1.ColumnCount-1 do
     FrmGrid1.cxGrid1DBTableView1.Columns[i].Width:=80;
  with FrmGrid1.cxGrid1DBTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDEXPERTIZE').Index].Caption:='Код экспертизы';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDEXPERTIZE').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NP').Index].Caption:='Номер протокола';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NP').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].Caption:='Индекс протокола';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('SPETSIALIST').Index].Caption:='Ф.И.О. специалиста';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEEXPERTIZE').Index].Caption:='Тип экспертизы';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEEXPERTIZE').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('CODEOBJ').Index].Caption:='Наименование объекта';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('CODEOBJ').Index].Width:=200;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('SOSTOYANIE').Index].Caption:='Состояние объекта';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('PRIMECH').Index].Caption:='Примечание объекта';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DATEEXPERTIZE').Index].Caption:='Дата экспертизы';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DATEPROTOKOL').Index].Caption:='Дата протокола';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('DATEPROTOKOL').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].Caption:='Клиент';
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDROW').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDROW').Index].Hidden:=true;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName(IDField).Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName(IDField).Index].Hidden:=true;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('IDREGISTRATION').Index].Hidden:=true;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('SPETSIALIST').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('SPETSIALIST').Index].Properties) do
        begin
          ListSource := FrmDialog2.DataSource1;
          ListFieldNames := 'NAMESPETSIALIST';
          KeyFieldNames := 'IDSPETSIALIST';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      ADODataSet2.Active:=false;
      ADODataSet2.CommandText:='select * from '+ObjTbl;
      ADODataSet2.Active:=true;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('CODEOBJ').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('CODEOBJ').Index].Properties) do
        begin
          ListSource := DataSource2;
          case FrmPeriodExp.RadioGroupObj.ItemIndex of
            0: begin
                ListFieldNames := 'NAMEOBJ;NAMEOBJRUS';
                KeyFieldNames := 'IDOBJ';
               end;
            1,2: begin
                ListFieldNames := 'NAMELAT;NAMERUS';
                KeyFieldNames := 'IDNEKAROBJ';
               end;
          end;
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ORGANIZATION').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource4;
          ListFieldNames := 'NAMECLIENT';
          KeyFieldNames := 'IDCLIENT';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
    end;
    FrmGrid1.ButtonDelete.Visible:=false;
    FrmGrid1.Show;
    FrmGrid1.ButtonAdd.DataSource:=nil;
    FrmGrid1.cxGrid1DBTableView1.OptionsData.Editing:=FullDostupBool;
    FrmGrid1.cxGrid1DBTableView1.OptionsData.Deleting:=false;
    FrmGrid1.cxGrid1DBTableView1.OptionsData.Appending:=FullDostupBool;
    //if FullDostupBool=false then FrmGrid1.ButtonDelete.DataSource:=nil;
    //FrmGrid1.ButtonDelete.DataSource:=nil;
  end; // если идёт выборка по коду ----------------------------------------------
end;

procedure TFrmMainWindow.AdvGlowButton2Click(Sender: TObject);
var
  i, n: integer;
  z, SQLstr, ina: string;
  d: variant;
begin
  if FrmDialog1= NIL then FrmDialog1 := TFrmDialog1.Create(owner);
  if ADOConnection1.Connected=false then ADOConnection1.Connected:=true;
  FrmPeriodOtch.ADODataSet1.Active:=false;
  FrmPeriodOtch.ADODataSet1.CommandText:='SELECT * FROM FIELDOTCH';
  FrmPeriodOtch.ADODataSet1.Active:=true;
  {for i:=0 to ADODataSet1.FieldCount-1 do
      begin
        FrmPeriodOtch.ADOCommand1.CommandText:='INSERT INTO FIELDOTCH ( NEED, FIELDNAME ) SELECT False AS N, "'+ADODataSet1.Fields[i].FieldName+'" AS FN';
        FrmPeriodOtch.ADOCommand1.Execute;
      end; }
  FrmPeriodOtch.cxGrid1DBTableView1.ClearItems;
  FrmPeriodOtch.cxGrid1DBTableView1.DataController.DataSource:=FrmPeriodOtch.DataSource1;
  with FrmPeriodOtch.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with FrmPeriodOtch.cxGrid1DBTableView1 do
    begin
      Columns[FrmPeriodOtch.cxGrid1DBTableView1.GetColumnByFieldName('NEED').Index].Caption:='Включить в отчёт';
      Columns[FrmPeriodOtch.cxGrid1DBTableView1.GetColumnByFieldName('NEED').Index].Width:=110;
      Columns[FrmPeriodOtch.cxGrid1DBTableView1.GetColumnByFieldName('FIELDNAME').Index].Visible:=false;
      Columns[FrmPeriodOtch.cxGrid1DBTableView1.GetColumnByFieldName('FIELDNAME').Index].Hidden:=true;
      Columns[FrmPeriodOtch.cxGrid1DBTableView1.GetColumnByFieldName('FIELDCAPTION').Index].Caption:='Название поля';
      Columns[FrmPeriodOtch.cxGrid1DBTableView1.GetColumnByFieldName('FIELDCAPTION').Index].Options.Editing:=false;
      Columns[FrmPeriodOtch.cxGrid1DBTableView1.GetColumnByFieldName('FIELDCAPTION').Index].Width:=320;
    end;
  FrmPeriodOtch.Showmodal;
  if PerStart='' then exit;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  ADODataSet1.Active:=false;
  ADODataSet1.CommandText:='SELECT FIELDNAME, FIELDCAPTION FROM FIELDOTCH WHERE (((NEED)=True))';
  ADODataSet1.Active:=true;
  if ADODataSet1.RecordCount=0 then
      begin
        showmessage('Не выбрано ни одно поле для отчёта!');
        exit;
      end;
  n:=ADODataSet1.RecordCount;
  d:= VarArrayCreate([1,n],VarOleStr);
  ADODataSet1.First;
  i:=0;
  if n>1 then for i:=1 to n do
    begin
      if i=1 then
        z:=ADODataSet1.Fields[0].AsString
      else
        z:=z+', '+ADODataSet1.Fields[0].AsString;
      d[i]:=ADODataSet1.Fields[1].AsString;
      ADODataSet1.Next;
    end;
  Screen.Cursor:=crHourGlass;
  //определяем доступ к отчётам
  ADODataSet1.Active:=false;
  SQLstr:='SELECT FULLDOSTUP, EXPVREDITEL, EXPGRIB, EXPBAKTER, EXPVIRUS, EXPNEMATOD, EXPSORN FROM USERS WHERE (((NAMECOMP)="'+GetComputerNetName+'"));';
  ADODataSet1.CommandText:=SQLstr;
  ADODataSet1.Active:=true;
  if ADODataSet1.FieldByName('EXPVREDITEL').AsBoolean=true then ina:='"Вредители растений"';
  if ADODataSet1.FieldByName('EXPGRIB').AsBoolean=true then if ina<>'' then ina:=ina+', "Болезни растений, грибные"' else ina:='"Болезни растений, грибные"';
  if ADODataSet1.FieldByName('EXPBAKTER').AsBoolean=true then if ina<>'' then ina:=ina+', "Болезни растений, бактериальные"' else ina:='"Болезни растений, бактериальные"';
  if ADODataSet1.FieldByName('EXPVIRUS').AsBoolean=true then if ina<>'' then ina:=ina+', "Болезни растений, вирусные"' else ina:='"Болезни растений, вирусные"';
  if ADODataSet1.FieldByName('EXPNEMATOD').AsBoolean=true then if ina<>'' then ina:=ina+', "Болезни растений, нематодные"' else ina:='"Болезни растений, нематодные"';
  if ADODataSet1.FieldByName('EXPSORN').AsBoolean=true then if ina<>'' then ina:=ina+', "Сорные растения"' else ina:='"Сорные растения"';
  if ShowNotFoundObj=true then if ina<>'' then ina:=ina+', "Не обнаружены"' else ina:='"Не обнаружены"';
  if (ina='') and (ADODataSet1.FieldByName('FULLDOSTUP').AsBoolean=false) then
     begin
        showmessage('Нет доступа для построения отчёта, обратитесь в организационный отдел!');
        exit;
      end;
  if ADODataSet1.FieldByName('FULLDOSTUP').AsBoolean=false then i:=-1;
  //формируем саму таблицу
  FrmGrid1.cxGrid1DBTableView1.ClearItems;
  ADODataSet1.Active:=false;
  SQLstr:='SELECT '+z+', "Карантинные" AS TYPEO FROM EXP_KAROBJ WHERE (((DATEPROTOKOL)>=#'+PerStart+'# And (DATEPROTOKOL)<=#'+PerFinish+'#))';
  SQLstr:=SQLstr+' union all select '+z+', "Некарантинные, отсутствующие в РФ" AS TYPEO from EXP_NEKAROTSRF WHERE (((DATEPROTOKOL)>=#'+PerStart+'# And (DATEPROTOKOL)<=#'+PerFinish+'#))';
  SQLstr:=SQLstr+' union all select '+z+', "Некарантинные, прочие" AS TYPEO from EXP_NEKARPROCH WHERE (((DATEPROTOKOL)>=#'+PerStart+'# And (DATEPROTOKOL)<=#'+PerFinish+'#))';
  if ShowNotFoundObj=true then SQLstr:=SQLstr+' union all select '+z+', "Не обнаружены" AS TYPEO from EXP_OBJNOTFOUND WHERE (((DATEPROTOKOL)>=#'+PerStart+'# And (DATEPROTOKOL)<=#'+PerFinish+'#))';
  if i=-1 then SQLstr:='select * from ('+SQLstr+') WHERE (((TYPEEXPERTIZE) In ('+ina+')))';
  ADODataSet1.CommandText:=SQLstr;
  try
    ADODataSet1.Active:=true;
  except
    Screen.Cursor:=crDefault;
    showmessage('Не возможно построить отчёт, попробуйте выбрать другую комбинацию полей!');
    exit;
  end;
  FrmGrid1.cxGrid1DBTableView1.DataController.DataSource:=DataSource1;
  with FrmGrid1.cxGrid1DBTableView1.DataController do
    CreateAllItems;
  for i:=1 to n do
    with FrmGrid1.cxGrid1DBTableView1 do
      Columns[i-1].Caption:=d[i];
  FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEO').Index].Caption:='Категория объекта';
  FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('TYPEO').Index].GroupBy(0,true);
  FrmGrid1.cxGrid1DBTableView1.ViewData.Expand(True);
  for i:=0 to FrmGrid1.cxGrid1DBTableView1.ColumnCount-1 do
     FrmGrid1.cxGrid1DBTableView1.Columns[i].Width:=80;
  try
    FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].PropertiesClass:=TcxImageComboBoxProperties;
    with FrmGrid1.cxGrid1DBTableView1 do
      begin
          TcxImageComboBoxProperties(FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Images:=FrmGrid1.ImageList2;
          TcxImageComboBoxProperties(FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items.Add;
          TcxImageComboBoxProperties(FrmGrid1.cxGrid1DBTableView1.Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items.Add;
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items[0].Description:='Импортный';
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items[0].Value:=0;
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items[1].Description:='Отечественный';
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items[1].Value:=1;
      end;
  except
  end;
  FrmGrid1.Show;
  FrmGrid1.ButtonAdd.DataSource:=nil;
  FrmGrid1.ButtonDelete.DataSource:=nil;
  FrmGrid1.cxGrid1DBTableView1.OptionsData.Editing:=false;
  FrmGrid1.cxGrid1DBTableView1.OptionsData.Deleting:=false;
  FrmGrid1.cxGrid1DBTableView1.OptionsData.Appending:=false;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.ButtonShowExcelClick(Sender: TObject);
begin
  try
    ExcelApp.DisplayFullScreen:=true;
    ExcelApp.DisplayFullScreen:=false;
  except
    ButtonShowExcel.Enabled:=false;
  end;
end;

procedure TFrmMainWindow.ButtonAboutClick(Sender: TObject);
begin
 if FrmAbout= NIL then FrmAbout := TFrmAbout.Create(owner);
 FrmAbout.Show;
end;

procedure TFrmMainWindow.ButtonNewRegClick(Sender: TObject);
var
  InputString: string;
begin
  InputString:= InputBox('Регистрация данного ПК в списке доступа:','Введите пароль', '*****');
  if InputString='A4tech' then
    begin
      FrmDialog1.ADOCommand1.CommandText:='INSERT INTO USERS (NAMECOMP, NAMEUSER, FULLDOSTUP) SELECT "'+GetComputerNetName+'" AS COMPN, "Безымянный пользователь" AS USERN, True AS DOSTUP';
      FrmDialog1.ADOCommand1.Execute;
    end
  else
    showmessage('Не правильный пароль!');
end;

procedure TFrmMainWindow.N21Click(Sender: TObject);
begin
  NekarObjProchie(N21.Caption);
end;

procedure TFrmMainWindow.N22Click(Sender: TObject);
begin
  NekarObjProchie(N22.Caption);
end;

procedure TFrmMainWindow.N26Click(Sender: TObject);
begin
  NekarObjProchie(N26.Caption);
end;

procedure TFrmMainWindow.N23Click(Sender: TObject);
begin
  NekarObjProchie(N23.Caption);
end;

procedure TFrmMainWindow.N24Click(Sender: TObject);
begin
  NekarObjProchie(N24.Caption);
end;

procedure TFrmMainWindow.N25Click(Sender: TObject);
begin
  NekarObjProchie(N25.Caption);
end;

procedure TFrmMainWindow.ButtonSvodTemplateClick(Sender: TObject);
begin
  ButtonShowExcel.Enabled:=true;
  try
    ExcelApp := GetActiveOleObject('Excel.Application');   //проверяем, нет ли запущенного Excel
  except
    on EOLESysError do ExcelApp:= CreateOleObject('Excel.Application');   //если нет, то запускаем
  end;
  ExcelApp.Application.EnableEvents:= false; // так будет быстрее
  ExcelApp.Application.DisplayAlerts:= true; // спрашивать в Excel сохранять ли файл после изменений
  ExcelApp.Visible:=true;
  Workbook := ExcelApp.WorkBooks.Open(ExtractFilePath(Application.ExeName)+'Templates\СводныйОтчёт1.xls');
  ExcelApp.DisplayFullScreen:=true;
  ExcelApp.DisplayFullScreen:=false;
end;

procedure TFrmMainWindow.ButtonSvodOtchClick(Sender: TObject);
var
  SQLstr, SQLa: string;
  i, pei, peo: integer;
  Excel : Variant;
begin
  FrmPeriod.Showmodal;
  if PerStart='' then exit;
  //проверяем, запущен ли Excel
  try
    ExcelApp := GetActiveOleObject('Excel.Application');
  except
    //если запущен, то "прикрываем"
    if VarIsEmpty(Excel) = false then
      begin
        Excel.Quit;
        Excel := 0;
      end;
  end;
  //открываем "наш" Excel
  try
    ExcelApp := CreateOleObject('Excel.Application');
  except
    ShowMessage('Ошибка при запуске MS Excel!');
    exit;
  end;
  ExcelApp.Application.EnableEvents:= false; // так будет быстрее
  ExcelApp.Application.DisplayAlerts:=false; // спрашивать в Excel сохранять ли файл после изменений
  ExcelApp.Visible:=true;
  Screen.Cursor:=crDefault;
  Workbook := ExcelApp.WorkBooks.Open(ExtractFilePath(Application.ExeName)+'Templates\СводныйОтчёт1.xls',0,true);
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[3, 4] := PerStart; //дата начальная
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[3, 7] := PerFinish; //и дата конечная
  //проанализировано образцов
  ADODataSet2.Active:=false;
  ADODataSet2.CommandText:='SELECT Sum(IIf([ISIMPORT]=0,[KOLOBR],0)) AS KOLIMP, Sum(IIf([ISIMPORT]=1,[KOLOBR],0)) AS KOLOTECH, Sum([REGISTRATION].[KOLOBR]) AS ITOGOOBR FROM REGISTRATION WHERE ((([REGISTRATION].[DATEPROTOKOL])>=#'+PerStart+'# And ([REGISTRATION].[DATEPROTOKOL])<=#'+PerFinish+'#));';
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[8, 1]:= ADODataSet2.FieldByName('KOLIMP').AsString;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[8, 2]:= ADODataSet2.FieldByName('KOLOTECH').AsString;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[8, 4]:= ADODataSet2.FieldByName('ITOGOOBR').AsString;
  //итого в таблицах
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[24, 3]:= ADODataSet2.FieldByName('KOLIMP').AsString;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[36, 3]:= ADODataSet2.FieldByName('KOLOTECH').AsString;
  //проведено экспертиз
  ADODataSet2.Active:=false;
  ADODataSet2.CommandText:='SELECT Sum(IIf([ISIMPORT]=0,1,0)) AS KOLIMP, Sum(IIf([ISIMPORT]=1,1,0)) AS KOLOTECH, Sum(1) AS ITOGEXP FROM EXPERTIZE_REESTR WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#));';
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[10, 1]:= ADODataSet2.FieldByName('KOLIMP').AsString;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[10, 2]:= ADODataSet2.FieldByName('KOLOTECH').AsString;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[10, 4]:= ADODataSet2.FieldByName('ITOGEXP').AsString;
  ADODataSet2.Active:=false;
  //в том числе из других областей
  ADODataSet2.Active:=false;
  SQLstr:='SELECT Sum(IIf([EXPERTIZE_REESTR]![ISIMPORT]=0,1,0)) AS KOLIMP, Sum(IIf([EXPERTIZE_REESTR]![ISIMPORT]=1,1,0)) AS KOLOTECH, Sum(1) AS ITOGEXP FROM EXPERTIZE_REESTR LEFT JOIN REGISTRATION ON EXPERTIZE_REESTR.NUMPROTOKOL = REGISTRATION.IDREGISTRATION';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((REGISTRATION.UPRSHNADZ) Not In (1,7)))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[12, 1]:= ADODataSet2.FieldByName('KOLIMP').AsString;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[12, 2]:= ADODataSet2.FieldByName('KOLOTECH').AsString;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[10, 4]:= ADODataSet2.FieldByName('ITOGEXP').AsString;
  ADODataSet2.Active:=false;

  //выявлено при карантинной экспертизе
  ADODataSet2.Active:=false;
  SQLa:='SELECT EXPERTIZE_REESTR.IDEXPERTIZE, First(REGISTRATION.KOLOBR) AS KOLANOBR, EXPERTIZE_REESTR.ISIMPORT, SPETSIALISTS_LAB.TYPESPETSIALIST AS TYPEEXPERTIZE';
  SQLa:=SQLa+' FROM ((EXPERTIZE_REESTR RIGHT JOIN REGISTRATION ON EXPERTIZE_REESTR.NUMPROTOKOL = REGISTRATION.IDREGISTRATION) LEFT JOIN EXPERTIZE_DATA ON EXPERTIZE_REESTR.IDEXPERTIZE = EXPERTIZE_DATA.NUMEXPERTIZE)';
  SQLa:=SQLa+' LEFT JOIN SPETSIALISTS_LAB ON EXPERTIZE_REESTR.SPETSIALIST = SPETSIALISTS_LAB.IDSPETSIALIST';
  SQLa:=SQLa+' WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#))';
  SQLa:=SQLa+' GROUP BY EXPERTIZE_REESTR.IDEXPERTIZE, EXPERTIZE_REESTR.ISIMPORT, SPETSIALISTS_LAB.TYPESPETSIALIST';
  pei:=0;
  peo:=0;
  //считаем количество по типам экспертиз
  SQLstr:='SELECT Sum(IIf([ISIMPORT]=1,[KOLANOBR],0)) AS KOLOTECH, Sum(IIf([ISIMPORT]=0,[KOLANOBR],0)) AS KOLIMP FROM ('+SQLa+') AS A WHERE ((([A].TYPEEXPERTIZE)="Вредители растений"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 3]:= ADODataSet2.FieldByName('KOLIMP').AsString;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[31, 3]:= ADODataSet2.FieldByName('KOLOTECH').AsString;
  pei:=pei+ADODataSet2.FieldByName('KOLIMP').AsInteger;
  peo:=peo+ADODataSet2.FieldByName('KOLOTECH').AsInteger;
  ADODataSet2.Active:=false;
  SQLstr:='SELECT Sum(IIf([ISIMPORT]=1,[KOLANOBR],0)) AS KOLOTECH, Sum(IIf([ISIMPORT]=0,[KOLANOBR],0)) AS KOLIMP FROM ('+SQLa+') AS A WHERE ((([A].TYPEEXPERTIZE)="Болезни растений, грибные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[20, 3]:= ADODataSet2.FieldByName('KOLIMP').AsString;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[32, 3]:= ADODataSet2.FieldByName('KOLOTECH').AsString;
  pei:=pei+ADODataSet2.FieldByName('KOLIMP').AsInteger;
  peo:=peo+ADODataSet2.FieldByName('KOLOTECH').AsInteger;
  ADODataSet2.Active:=false;
  SQLstr:='SELECT Sum(IIf([ISIMPORT]=1,[KOLANOBR],0)) AS KOLOTECH, Sum(IIf([ISIMPORT]=0,[KOLANOBR],0)) AS KOLIMP FROM ('+SQLa+') AS A WHERE ((([A].TYPEEXPERTIZE)="Болезни растений, нематодные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[21, 3]:= ADODataSet2.FieldByName('KOLIMP').AsString;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[33, 3]:= ADODataSet2.FieldByName('KOLOTECH').AsString;
  pei:=pei+ADODataSet2.FieldByName('KOLIMP').AsInteger;
  peo:=peo+ADODataSet2.FieldByName('KOLOTECH').AsInteger;
  ADODataSet2.Active:=false;
  SQLstr:='SELECT Sum(IIf([ISIMPORT]=1,[KOLANOBR],0)) AS KOLOTECH, Sum(IIf([ISIMPORT]=0,[KOLANOBR],0)) AS KOLIMP FROM ('+SQLa+') AS A WHERE ((([A].TYPEEXPERTIZE)="Болезни растений, бактериальные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[22, 3]:= ADODataSet2.FieldByName('KOLIMP').AsString;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[34, 3]:= ADODataSet2.FieldByName('KOLOTECH').AsString;
  pei:=pei+ADODataSet2.FieldByName('KOLIMP').AsInteger;
  peo:=peo+ADODataSet2.FieldByName('KOLOTECH').AsInteger;
  ADODataSet2.Active:=false;
  SQLstr:='SELECT Sum(IIf([ISIMPORT]=1,[KOLANOBR],0)) AS KOLOTECH, Sum(IIf([ISIMPORT]=0,[KOLANOBR],0)) AS KOLIMP FROM ('+SQLa+') AS A WHERE ((([A].TYPEEXPERTIZE)="Сорные растения"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[23, 3]:= ADODataSet2.FieldByName('KOLIMP').AsString;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[35, 3]:= ADODataSet2.FieldByName('KOLOTECH').AsString;
  pei:=pei+ADODataSet2.FieldByName('KOLIMP').AsInteger;
  peo:=peo+ADODataSet2.FieldByName('KOLOTECH').AsInteger;
  //если не было по каким-то экспертиз вообще, то забиваем нули
  for i:=19 to 24 do if ExcelApp.WorkBooks[1].WorkSheets[1].Cells[i, 3].Text='' then ExcelApp.WorkBooks[1].WorkSheets[1].Cells[i, 3]:='0';
  for i:=31 to 36 do if ExcelApp.WorkBooks[1].WorkSheets[1].Cells[i, 3].Text='' then ExcelApp.WorkBooks[1].WorkSheets[1].Cells[i, 3]:='0';
  //итого выявлено при карантинной экспертизе импортных подкарантинных материалов
  {ADODataSet2.Active:=false;
  ADODataSet2.CommandText:='SELECT EXPERTIZE_REESTR.NUMPROTOKOL FROM EXPERTIZE_REESTR WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#6/1/2007# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#2/1/2008#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0)) GROUP BY EXPERTIZE_REESTR.NUMPROTOKOL;';
  ADODataSet2.Active:=true;  }
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[10, 1]:=pei;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[10, 2]:=peo;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[10, 4]:=pei+peo;

  //обнаружено видов импортных
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Вредители растений"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=0))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 5]:= IntToStr(ADODataSet2.RecordCount); //виды карантинные импортные
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Вредители растений"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=1))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 9]:= IntToStr(ADODataSet2.RecordCount);  //виды некар отсутств имп
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Вредители растений"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=2))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 7]:= IntToStr(ADODataSet2.RecordCount);  //виды некар прочие

  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, грибные"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=0))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[20, 5]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, грибные"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=1))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[20, 9]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, грибные"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=2))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[20, 7]:= IntToStr(ADODataSet2.RecordCount);

  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, нематодные"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=0))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[21, 5]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, нематодные"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=1))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[21, 9]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, нематодные"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=2))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[21, 7]:= IntToStr(ADODataSet2.RecordCount);

  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, бактериальные"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=2))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[22, 5]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, бактериальные"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=1))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[22, 9]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, бактериальные"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=2))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[22, 7]:= IntToStr(ADODataSet2.RecordCount);

  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Сорные растения"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=0))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[23, 5]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Сорные растения"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=1))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[23, 9]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Сорные растения"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=2))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[23, 7]:= IntToStr(ADODataSet2.RecordCount);

  //обнаружено видов отечественных
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Вредители растений"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=0))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[31, 5]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Вредители растений"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=1))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[31, 9]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Вредители растений"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=2))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[31, 7]:= IntToStr(ADODataSet2.RecordCount);

  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, грибные"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=0))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[32, 5]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, грибные"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=1))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[32, 9]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, грибные"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=2))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[32, 7]:= IntToStr(ADODataSet2.RecordCount);

  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, нематодные"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=0))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[33, 5]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, нематодные"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=1))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[33, 9]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, нематодные"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=2))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[33, 7]:= IntToStr(ADODataSet2.RecordCount);

  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, бактериальные"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=0))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[34, 5]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, бактериальные"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=1))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[34, 9]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, бактериальные"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=2))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[34, 7]:= IntToStr(ADODataSet2.RecordCount);

  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Сорные растения"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=0))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[35, 5]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Сорные растения"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=1))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[35, 9]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+'  WHERE (((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Сорные растения"))';
  SQLstr:=SQLstr+'  GROUP BY EXPERTIZE_DATA.CODEOBJ, EXPERTIZE_DATA.TYPEOBJ HAVING (((EXPERTIZE_DATA.TYPEOBJ)=2))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[35, 7]:= IntToStr(ADODataSet2.RecordCount);

  //случаев импортные
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=0) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Вредители растений"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 6]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=1) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Вредители растений"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 10]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=2) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Вредители растений"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 8]:= IntToStr(ADODataSet2.RecordCount);

  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=0) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, грибные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[20, 6]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=1) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, грибные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[20, 10]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=2) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, грибные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[20, 8]:= IntToStr(ADODataSet2.RecordCount);

  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=0) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, нематодные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[21, 6]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=1) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, нематодные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[21, 10]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=2) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, нематодные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[21, 8]:= IntToStr(ADODataSet2.RecordCount);

  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=0) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, бактериальные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[22, 6]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=1) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, бактериальные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[22, 10]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=2) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, бактериальные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[22, 8]:= IntToStr(ADODataSet2.RecordCount);

  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=0) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Сорные растения"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[23, 6]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=1) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Сорные растения"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[23, 10]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=2) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=0) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Сорные растения"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[23, 8]:= IntToStr(ADODataSet2.RecordCount);

  //случаев отечественные
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=0) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Вредители растений"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[31, 6]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=1) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Вредители растений"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[31, 10]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=2) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Вредители растений"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[31, 8]:= IntToStr(ADODataSet2.RecordCount);

  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=0) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, грибные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[32, 6]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=1) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, грибные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[32, 10]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=2) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, грибные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[32, 8]:= IntToStr(ADODataSet2.RecordCount);

  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=0) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, нематодные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[33, 6]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=1) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, нематодные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[33, 10]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=2) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, нематодные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[33, 8]:= IntToStr(ADODataSet2.RecordCount);

  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=0) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, бактериальные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[34, 6]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=1) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, бактериальные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[34, 10]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=2) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Болезни растений, бактериальные"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[34, 8]:= IntToStr(ADODataSet2.RecordCount);

  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=0) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Сорные растения"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[35, 6]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=1) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Сорные растения"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[35, 10]:= IntToStr(ADODataSet2.RecordCount);
  ADODataSet2.Active:=false;
  SQLstr:='SELECT EXPERTIZE_DATA.CODEOBJ FROM EXPERTIZE_DATA LEFT JOIN EXPERTIZE_REESTR ON EXPERTIZE_DATA.NUMEXPERTIZE = EXPERTIZE_REESTR.IDEXPERTIZE';
  SQLstr:=SQLstr+' WHERE (((EXPERTIZE_DATA.TYPEOBJ)=2) AND ((EXPERTIZE_REESTR.DATEEXPERTIZE)>=#'+PerStart+'# And (EXPERTIZE_REESTR.DATEEXPERTIZE)<=#'+PerFinish+'#) AND ((EXPERTIZE_REESTR.ISIMPORT)=1) AND ((EXPERTIZE_DATA.TYPEEXPERTIZE)="Сорные растения"))';
  ADODataSet2.CommandText:=SQLstr;
  ADODataSet2.Active:=true;
  ExcelApp.WorkBooks[1].WorkSheets[1].Cells[35, 8]:= IntToStr(ADODataSet2.RecordCount);

  //защищаем лист от изменений
  //ExcelApp.WorkBooks[1].WorkSheets[1].protect('344',true,true,false);
  ExcelApp.DisplayFullScreen:=true;
  ExcelApp.DisplayFullScreen:=false;
end;

procedure TFrmMainWindow.ButtonSvvo2TemplateClick(Sender: TObject);
begin
  ButtonShowExcel.Enabled:=true;
  try
    ExcelApp := GetActiveOleObject('Excel.Application');   //проверяем, нет ли запущенного Excel
  except
    on EOLESysError do ExcelApp:= CreateOleObject('Excel.Application');   //если нет, то запускаем
  end;
  ExcelApp.Application.EnableEvents:= false; // так будет быстрее
  ExcelApp.Application.DisplayAlerts:= true; // спрашивать в Excel сохранять ли файл после изменений
  ExcelApp.Visible:=true;
  Workbook := ExcelApp.WorkBooks.Open(ExtractFilePath(Application.ExeName)+'Templates\Свидетельство2.xls');
  ExcelApp.DisplayFullScreen:=true;
  ExcelApp.DisplayFullScreen:=false;
end;

procedure TFrmMainWindow.ButtonAktTemplateClick(Sender: TObject);
begin
  ButtonShowExcel.Enabled:=true;
  try
    ExcelApp := GetActiveOleObject('Excel.Application');   //проверяем, нет ли запущенного Excel
  except
    on EOLESysError do ExcelApp:= CreateOleObject('Excel.Application');   //если нет, то запускаем
  end;
  ExcelApp.Application.EnableEvents:= false; // так будет быстрее
  ExcelApp.Application.DisplayAlerts:= true; // спрашивать в Excel сохранять ли файл после изменений
  ExcelApp.Visible:=true;
  Workbook := ExcelApp.WorkBooks.Open(ExtractFilePath(Application.ExeName)+'Templates\Акт.xls');
  ExcelApp.DisplayFullScreen:=true;
  ExcelApp.DisplayFullScreen:=false;
end;

procedure TFrmMainWindow.AdvGlowButton5Click(Sender: TObject);
begin
  if FrmToolExcel= NIL then FrmToolExcel := TFrmToolExcel.Create(owner);
  FrmToolExcel.Show;
end;

procedure TFrmMainWindow.Timer1Timer(Sender: TObject);
  procedure Delay(msec: Longint);
   var
     start, stop: Longint;
   begin
     start := GetTickCount;
     repeat
       stop := GetTickCount;
       Application.ProcessMessages;
     until (stop - start) >= msec;
   end;
var
   maxx, maxy: Integer;
begin
   {fc:=fc+1;
   if fc=1 then exit;
   maxx := ImageZnak.Width;
   maxy := ImageZnak.Height;
   ImageZnak.Width  := 112;
   ImageZnak.Height := 27;
   ImageZnak.Left   := (Width - ImageZnak.Width) div 2;
   ImageZnak.Top    := (Height - ImageZnak.Height) div 2;
   ImageZnak.Show;
   repeat
     if ImageZnak.Height + (maxy div 5) >= maxy then
       ImageZnak.Height := maxy
     else
       ImageZnak.Height := ImageZnak.Height + (maxy div 5);
     if ImageZnak.Width + (maxx div 5) >= maxx then
       ImageZnak.Width := maxx
     else
       ImageZnak.Width := ImageZnak.Width + (maxx div 5);
     ImageZnak.Left := (Width - ImageZnak.Width) div 2;
     ImageZnak.Top  := ((Height - ImageZnak.Height) div 2)+40;
     delay(30);
   until (ImageZnak.Width = maxx) and (ImageZnak.Height = maxy);
  Timer1.Enabled:=false; }
  WiiProgressBar1.Left:=Width-75;
  WiiProgressBar1.Top:=Height-250;
end;

procedure TFrmMainWindow.ButtonZhurnalClick(Sender: TObject);
var
  i:integer;
  SQLstr: string;
begin
  FrmPeriod.Showmodal;
  if PerStart='' then exit;
  Screen.Cursor:=crHourGlass;
  if FrmGrid1= NIL then FrmGrid1 := TFrmGrid1.Create(owner);
  if FrmDialog1= NIL then FrmDialog1 := TFrmDialog1.Create(owner);
  if ADOConnection1.Connected=false then ADOConnection1.Connected:=true;
  FrmGrid1.cxGrid1Level1.GridView:=FrmGrid1.cxGrid1DBBandedTableView1;
  //обновление источников данных
  FrmDialog1.ADODataSet1.Active:=false;
  FrmDialog1.ADODataSet2.Active:=false;
  FrmDialog1.ADODataSet3.Active:=false;
  FrmDialog1.ADODataSet4.Active:=false;
  FrmDialog1.ADODataSet5.Active:=false;
  FrmDialog1.ADODataSet6.Active:=false;
  FrmDialog1.ADODataSet7.Active:=false;
  FrmDialog1.ADODataSet8.Active:=false;
  FrmDialog1.ADODataSet9.Active:=false;
  FrmDialog1.ADODataSet10.Active:=false;
  FrmDialog1.ADODataSet11.Active:=false;
  FrmDialog1.ADODataSet13.Active:=false;
  FrmDialog1.ADODataSet1.Active:=true;
  FrmDialog1.ADODataSet2.Active:=true;
  FrmDialog1.ADODataSet3.Active:=true;
  FrmDialog1.ADODataSet4.Active:=true;
  FrmDialog1.ADODataSet5.Active:=true;
  FrmDialog1.ADODataSet6.Active:=true;
  FrmDialog1.ADODataSet7.Active:=true;
  FrmDialog1.ADODataSet8.Active:=true;
  FrmDialog1.ADODataSet9.Active:=true;
  FrmDialog1.ADODataSet10.Active:=true;
  FrmDialog1.ADODataSet11.Active:=true;
  FrmDialog1.ADODataSet13.Active:=true;
  ADODataSet1.Active:=false;
  FrmGrid1.cxGrid1DBBandedTableView1.Bands.Clear;
  FrmGrid1.cxGrid1DBBandedTableView1.ClearItems;
  SQLstr:='SELECT IDREGISTRATION, ISIMPORT, NUMPROTOKOL, DATEPROTOKOL, DOСTYPE, UPRSHNADZ, PUNKTKARRAST, SPETSIALIST, IIf([NUMETIKETKI]<>"0",[NUMETIKETKI],"") AS NUMETIKETKI, IIf([NUMETIKETKI]<>"0",[ETIKETKADATE],"") AS ETIKETKADATE, ORGANIZATION,';
  SQLstr:=SQLstr+' NUMDOGOVOR, IIf([NUMZAYAVKA]<>"0",[NUMZAYAVKA],"") AS NUMZAYAVKA,';
  SQLstr:=SQLstr+' IIf([NUMZAYAVKA]<>"0",[DATEZAYAVKA],"") AS DATEZAYAVKA, NAMEOBRAZTSA, KLASSIFIKATION, NAMEPATRY, IIf([SORT]<>"нет",[SORT],"") AS SORT,';
  SQLstr:=SQLstr+' IIf([ISIMPORT]=0,[PROISHOZHD],0) AS PROISHOZHD, IIf([VES]<>0,[VES],"") AS VES, EDVES, IIf([KOL]<>0,[KOL],"") AS KOL, EDIZM, IIf([KOLOBR]<>0,[KOLOBR],"") AS KOLOBR, IIf([SHTUKVOBR]<>0,[SHTUKVOBR],"") AS SHTUKVOBR,';
  SQLstr:=SQLstr+'  VIDTRANSPORT, NAZVTRANSPORT, PUNKTNAZNACH, VUPRSHNADZ, VCLIENT, VARCHIV, PUPRSHNADZ, PCLIENT, PREFCENTR, EXPCLOSE, OTKAZKL FROM REGISTRATION WHERE (((DATEPROTOKOL)>=#'+PerStart+'# And (DATEPROTOKOL)<=#'+PerFinish+'#))';
  ADODataSet1.CommandText:=SQLstr;
  ADODataSet1.Active:=true;
  FrmGrid1.cxGrid1DBBandedTableView1.DataController.DataSource:=DataSource1;
  FrmGrid1.cxGrid1DBBandedTableView1.Bands.Add;
  FrmGrid1.cxGrid1DBBandedTableView1.Bands.Add;
  FrmGrid1.cxGrid1DBBandedTableView1.Bands.Add;
  FrmGrid1.cxGrid1DBBandedTableView1.Bands.Add;
  FrmGrid1.cxGrid1DBBandedTableView1.Bands.Add;
  FrmGrid1.cxGrid1DBBandedTableView1.Bands.Add;
  FrmGrid1.cxGrid1DBBandedTableView1.Bands.Add;
  FrmGrid1.cxGrid1DBBandedTableView1.Bands.Add;
  FrmGrid1.cxGrid1DBBandedTableView1.Bands.Add;
  FrmGrid1.cxGrid1DBBandedTableView1.Bands[0].Caption:='Протокол';
  FrmGrid1.cxGrid1DBBandedTableView1.Bands[1].Caption:='Сопроводительные документы';
  FrmGrid1.cxGrid1DBBandedTableView1.Bands[2].Position.BandIndex:=1;
  FrmGrid1.cxGrid1DBBandedTableView1.Bands[3].Position.BandIndex:=1;
  FrmGrid1.cxGrid1DBBandedTableView1.Bands[2].Caption:='Россельхознадзора (референтного центра)';
  FrmGrid1.cxGrid1DBBandedTableView1.Bands[3].Caption:='Клиента';
  FrmGrid1.cxGrid1DBBandedTableView1.Bands[4].Caption:='Наименование материала';
  FrmGrid1.cxGrid1DBBandedTableView1.Bands[5].Caption:='Кол-во обр.';
  FrmGrid1.cxGrid1DBBandedTableView1.Bands[6].Caption:='Штук в обр.';
  FrmGrid1.cxGrid1DBBandedTableView1.Bands[7].Caption:='Происхожд. и откуда поступил материал';
  FrmGrid1.cxGrid1DBBandedTableView1.Bands[8].Caption:='Пункт назначения'; //раньше был 'Пункт назначения (страна, регион, клиент)'
  with FrmGrid1.cxGrid1DBBandedTableView1.DataController do
    CreateAllItems;
  with FrmGrid1.cxGrid1DBBandedTableView1 do
    begin
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('IDREGISTRATION').Index].Caption:='Номер записи';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('ISIMPORT').Index].Caption:='Признак происхождения образца';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].Caption:='Номер';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('DATEPROTOKOL').Index].Caption:='Дата';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('DOСTYPE').Index].Caption:='Тип документа';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('UPRSHNADZ').Index].Caption:='Наименование Управления Россельхознадзора';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('PUNKTKARRAST').Index].Caption:='Название пункта по карантину растений';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('SPETSIALIST').Index].Caption:='Ф.И.О. специалиста, выдавшего этикетку';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NUMETIKETKI').Index].Caption:='Номер этикетки';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('ETIKETKADATE').Index].Caption:='Дата выдачи этикетки';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('ORGANIZATION').Index].Caption:='Название организации (Ф.И.О. частного лица)';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NUMDOGOVOR').Index].Caption:='Номер договора';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NUMZAYAVKA').Index].Caption:='Номер заявки клиента';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('DATEZAYAVKA').Index].Caption:='Дата заявки клиента';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NAMEOBRAZTSA').Index].Caption:='Наименование образца';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('KLASSIFIKATION').Index].Caption:='Классификация';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NAMEPATRY').Index].Caption:='Наим. партии продукции, от которой отобран обр.';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('SORT').Index].Caption:='Сорт образца';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('PROISHOZHD').Index].Caption:='Происхождение образца';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('VES').Index].Caption:='Вес';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('EDVES').Index].Caption:='Ед.изм. веса';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('KOL').Index].Caption:='Количество мест';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('EDIZM').Index].Caption:='Ед.изм. мест';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('KOLOBR').Index].Caption:='';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('SHTUKVOBR').Index].Caption:='';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('VIDTRANSPORT').Index].Caption:='Наим. трансп. средства';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NAZVTRANSPORT').Index].Caption:='Назв. трансп. ср-ва, номер, приход';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('PUNKTNAZNACH').Index].Caption:='';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('VUPRSHNADZ').Index].Caption:=' Управлению Россельхознадзора выдано Свидетельство карантинной экспертизы ';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('VCLIENT').Index].Caption:='Клиенту выдано Свидетельство карантинной экспертизы ';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('VARCHIV').Index].Caption:='В архиве ФГУ находится Свидетельство карантинной экспертизы ';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('PUPRSHNADZ').Index].Caption:='Из Управления Россельхознадзора поступили образцы';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('PCLIENT').Index].Caption:='От клиента поступили образцы';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('PREFCENTR').Index].Caption:='От Референтного центра поступили образцы';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('EXPCLOSE').Index].Caption:='Свидетельство выдано';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('OTKAZKL').Index].Caption:='Отказ клиента от анализа';
      //не нужные в журнале
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('ISIMPORT').Index].GroupBy(0,false);
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('IDREGISTRATION').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NUMDOGOVOR').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('KLASSIFIKATION').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('VUPRSHNADZ').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('VCLIENT').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('VARCHIV').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('PUPRSHNADZ').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('PCLIENT').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('PREFCENTR').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('EXPCLOSE').Index].Visible:=false;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('OTKAZKL').Index].Visible:=false;
      //группы столбцов
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].Position.BandIndex:=0;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('DATEPROTOKOL').Index].Position.BandIndex:=0;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('DOСTYPE').Index].Position.BandIndex:=2;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('UPRSHNADZ').Index].Position.BandIndex:=2;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('PUNKTKARRAST').Index].Position.BandIndex:=2;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('SPETSIALIST').Index].Position.BandIndex:=2;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NUMETIKETKI').Index].Position.BandIndex:=2;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('ETIKETKADATE').Index].Position.BandIndex:=2;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('ORGANIZATION').Index].Position.BandIndex:=3;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NUMZAYAVKA').Index].Position.BandIndex:=3;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('DATEZAYAVKA').Index].Position.BandIndex:=3;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NAMEOBRAZTSA').Index].Position.BandIndex:=4;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NAMEPATRY').Index].Position.BandIndex:=4;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('SORT').Index].Position.BandIndex:=4;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('PROISHOZHD').Index].Position.BandIndex:=4;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('VES').Index].Position.BandIndex:=4;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('EDVES').Index].Position.BandIndex:=4;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('KOL').Index].Position.BandIndex:=4;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('EDIZM').Index].Position.BandIndex:=4;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('KOLOBR').Index].Position.BandIndex:=5;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('SHTUKVOBR').Index].Position.BandIndex:=6;
      //Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('').Index].Position.BandIndex:=5;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('VIDTRANSPORT').Index].Position.BandIndex:=7;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NAZVTRANSPORT').Index].Position.BandIndex:=7;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('PUNKTNAZNACH').Index].Position.BandIndex:=8;
      //свойства столбцов
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('IDREGISTRATION').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NUMPROTOKOL').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('DATEPROTOKOL').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('ISIMPORT').Index].PropertiesClass:=TcxImageComboBoxProperties;
      with FrmGrid1.cxGrid1DBBandedTableView1 do
        begin
          TcxImageComboBoxProperties(FrmGrid1.cxGrid1DBBandedTableView1.Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Images:=FrmGrid1.ImageList2;
          TcxImageComboBoxProperties(FrmGrid1.cxGrid1DBBandedTableView1.Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items.Add;
          TcxImageComboBoxProperties(FrmGrid1.cxGrid1DBBandedTableView1.Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items.Add;
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items[0].Description:='Импортный';
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items[0].Value:=0;
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items[1].Description:='Отечественный';
          TcxImageComboBoxProperties(Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('ISIMPORT').Index].Properties).Items[1].Value:=1;
        end;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('ISIMPORT').Index].Options.Editing:=false;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('DOСTYPE').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('DOСTYPE').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource13;
          ListFieldNames := 'DOCSOKR';
          KeyFieldNames := 'IDDOCVIDS';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('DOСTYPE').Index].DataBinding.FieldName := 'DOСTYPE';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('UPRSHNADZ').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('UPRSHNADZ').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource3;
          ListFieldNames := 'NAMESOKR';
          KeyFieldNames := 'IDUPRSHNADZ';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('UPRSHNADZ').Index].DataBinding.FieldName := 'UPRSHNADZ';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('PUNKTKARRAST').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('PUNKTKARRAST').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource2;
          ListFieldNames := 'NAMEKARPUNKT';
          KeyFieldNames := 'IDKARPUNKT';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('PUNKTKARRAST').Index].DataBinding.FieldName := 'PUNKTKARRAST';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('SPETSIALIST').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('SPETSIALIST').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource1;
          ListFieldNames := 'NAMESPETSIALIST';
          KeyFieldNames := 'IDSPETSIALIST';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('SPETSIALIST').Index].DataBinding.FieldName := 'SPETSIALIST';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('ORGANIZATION').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('ORGANIZATION').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource4;
          ListFieldNames := 'NAMECLIENT';
          KeyFieldNames := 'IDCLIENT';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('ORGANIZATION').Index].DataBinding.FieldName := 'ORGANIZATION';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NAMEOBRAZTSA').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NAMEOBRAZTSA').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource5;
          ListFieldNames := 'NAMEOBR';
          KeyFieldNames := 'IDOBR';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NAMEOBRAZTSA').Index].DataBinding.FieldName := 'NAMEOBRAZTSA';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('KLASSIFIKATION').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('KLASSIFIKATION').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource7;
          ListFieldNames := 'NAMEKLASS';
          KeyFieldNames := 'IDKLASS';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('KLASSIFIKATION').Index].DataBinding.FieldName := 'KLASSIFIKATION';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NAMEPATRY').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NAMEPATRY').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource6;
          ListFieldNames := 'NAMEPARTII';
          KeyFieldNames := 'IDPARTII';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('NAMEPATRY').Index].DataBinding.FieldName := 'NAMEPATRY';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('PROISHOZHD').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('PROISHOZHD').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource8;
          ListFieldNames := 'NAMEPROISH';
          KeyFieldNames := 'IDPROISH';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('PROISHOZHD').Index].DataBinding.FieldName := 'PROISHOZHD';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('EDVES').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('EDVES').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource9;
          ListFieldNames := 'NAMEMEASURE';
          KeyFieldNames := 'IDMEASURE';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('EDVES').Index].DataBinding.FieldName := 'EDVES';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('EDIZM').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('EDIZM').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource10;
          ListFieldNames := 'NAMEEDIZM';
          KeyFieldNames := 'IDEDIZM';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('EDIZM').Index].DataBinding.FieldName := 'EDIZM';
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('VIDTRANSPORT').Index].PropertiesClass:=TcxLookupComboBoxProperties;
      with TcxLookupComboBoxProperties(Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('VIDTRANSPORT').Index].Properties) do
        begin
          ListSource := FrmDialog1.DataSource11;
          ListFieldNames := 'NAMETRANSPORT';
          KeyFieldNames := 'IDTRANSPORT';
          ImmediatePost:=true;
          ListOptions.ShowHeader:=false;
        end;
      Columns[FrmGrid1.cxGrid1DBBandedTableView1.GetColumnByFieldName('VIDTRANSPORT').Index].DataBinding.FieldName := 'VIDTRANSPORT';
    end;
  for i:=0 to FrmGrid1.cxGrid1DBBandedTableView1.ColumnCount-1 do
     FrmGrid1.cxGrid1DBBandedTableView1.Columns[i].Width:=80;
  FrmGrid1.cxGrid1DBBandedTableView1.ViewData.Expand(true);
  FrmGrid1.Show;
  FrmGrid1.ButtonAdd.DataSource:=nil;
  FrmGrid1.cxGrid1DBBandedTableView1.OptionsData.Editing:=FullDostupBool;
  FrmGrid1.cxGrid1DBBandedTableView1.OptionsData.Deleting:=FullDostupBool;
  FrmGrid1.cxGrid1DBBandedTableView1.OptionsData.Appending:=FullDostupBool;
  if FullDostupBool=false then FrmGrid1.ButtonDelete.DataSource:=nil;
  FrmGrid1.ToolBarExpertize.Visible:=false;
  FrmGrid1.ButtonGraph.Visible:=false;
  FrmGrid1.ButtonSwap.Visible:=false;
  Screen.Cursor:=crDefault;
end;

procedure TFrmMainWindow.ButtonSvvo3TemplateClick(Sender: TObject);
begin
  ButtonShowExcel.Enabled:=true;
  try
    ExcelApp := GetActiveOleObject('Excel.Application');   //проверяем, нет ли запущенного Excel
  except
    on EOLESysError do ExcelApp:= CreateOleObject('Excel.Application');   //если нет, то запускаем
  end;
  ExcelApp.Application.EnableEvents:= false; // так будет быстрее
  ExcelApp.Application.DisplayAlerts:= true; // спрашивать в Excel сохранять ли файл после изменений
  ExcelApp.Visible:=true;
  Workbook := ExcelApp.WorkBooks.Open(ExtractFilePath(Application.ExeName)+'Templates\Свидетельство3.xls');
  ExcelApp.DisplayFullScreen:=true;
  ExcelApp.DisplayFullScreen:=false;
end;

procedure TFrmMainWindow.ButtonUnblockRegProtokolClick(Sender: TObject);
begin
  ADOCommandLockRegistration.CommandText:='UPDATE SETUP SET SETUP.IDVAL = "0" WHERE (((SETUP.IDSET)=1))';
  ADOCommandLockRegistration.Execute;
end;

procedure TFrmMainWindow.AdvShapeButton2Click(Sender: TObject);
begin
   PreviewMenu1.ShowMenu(Left+5,Top+10);
end;

procedure TFrmMainWindow.PreviewMenu1Buttons0Click(Sender: TObject;
  Button: TButtonCollectionItem);
begin
  close;
end;

procedure TFrmMainWindow.PreviewMenu1MenuItems0Click(Sender: TObject);
begin
   if FrmOptions= NIL then FrmOptions := TFrmOptions.Create(owner);
   FrmOptions.ShowModal;
end;

procedure TFrmMainWindow.PreviewMenu1MenuItems2Click(Sender: TObject);
begin
   if FrmCalendar= NIL then FrmCalendar := TFrmCalendar.Create(owner);
   FrmCalendar.Show;
end;

procedure TFrmMainWindow.PreviewMenu1MenuItems1Click(Sender: TObject);
var
   SQLstr, t : string;
   i, v, z   : integer;
begin
   FrmPeriod.ShowModal;
   if PerStart='' then exit;
   Screen.Cursor:=crHourGlass;
   //спрашиваем про тип происхождения
   if PMMessageDialog('Тип происхождения'+Chr(13)+'выберите из двух вариантов', 'ВНИМАНИЕ:', mtWarning, mbOKCancel, ['Импорт', 'Отеч.']) = mrOk then t:='0' else t:='1';
   if FrmNumProtocols= NIL then FrmNumProtocols := TFrmNumProtocols.Create(owner);
   SQLstr:='SELECT REGISTRATION.NUMPROTOKOL FROM REGISTRATION' + #13#10 +
           'WHERE (((REGISTRATION.ISIMPORT)='+t+') AND ((REGISTRATION.DATEPROTOKOL)>=#'+PerStart+'# And (REGISTRATION.DATEPROTOKOL)<=#'+PerFinish+'#))' + #13#10 +
           'ORDER BY REGISTRATION.NUMPROTOKOL';
   FrmNumProtocols.ADODataSet2.Active:=false;
   FrmNumProtocols.ADODataSet2.CommandText:=SQLstr;
   FrmNumProtocols.ADODataSet2.Active:=true;
   FrmNumProtocols.cxGrid1DBTableView1.DataController.CreateAllItems(true);
   FrmNumProtocols.cxGrid1DBTableView1.Columns[0].Caption:='№ протокола';
   FrmNumProtocols.cxGrid1DBTableView1.Columns[0].Width:=100;
   try
      FrmNumProtocols.ADODataSet2.First;
      v:=FrmNumProtocols.ADODataSet2.FieldValues['NUMPROTOKOL'];
   except
      showmessage('Непредвиденная ошибка при попытке запроса номеров протоколов.');
      Screen.Cursor:=crDefault;
      exit;
   end;
   z:=v;
   for i:=1 to FrmNumProtocols.ADODataSet2.RecordCount-1 do
     begin
        if z<>FrmNumProtocols.ADODataSet2.FieldValues['NUMPROTOKOL'] then
           begin
              showmessage('Не найден протокол №'+IntToStr(z));
              Screen.Cursor:=crDefault;
              exit;
           end;
        z:=z+1;
        FrmNumProtocols.ADODataSet2.Next;
     end;
   Screen.Cursor:=crDefault;
   FrmNumProtocols.Show;
   showmessage('За данный период,'+Chr(13)+'по выбранному типу происхождения'+Chr(13)+'сохранена целостность нумерации протоколов!');
end;

procedure TFrmMainWindow.ButtonNekarObjOtsutRFClick(Sender: TObject);
begin
  PopupMenu1.Popup(Left+AdvToolBar10.Left+13,Top+120);
end;

procedure TFrmMainWindow.Timer2Timer(Sender: TObject);
var
  Db, De              : string;
  LYear, LMonth, LDay : Word;
begin
  if ViewStatistic=false then
    begin
      Timer2.Enabled:=false;
      WiiProgressBar1.Visible:=false;
      exit;
    end;  
  //статистика, всего
  Memo_statistics.Lines.Clear;
    ADODataSet_Statistics.Active:=false;
  ADODataSet_Statistics.CommandText:='SELECT Sum(KOLOBR) AS R FROM REGISTRATION;';
  ADODataSet_Statistics.Active:=true;
  if ADODataSet_Statistics.RecordCount>0 then
     Memo_statistics.Lines.Add('Зарегистрировано образцов, всего: '+ADODataSet_Statistics.FieldByName('R').AsString);
  ADODataSet_Statistics.Active:=false;
  Memo_statistics.Lines.Add('');
  ADODataSet_Statistics.Active:=false;
  ADODataSet_Statistics.CommandText:='SELECT Count(*) AS R FROM REGISTRATION;';
  ADODataSet_Statistics.Active:=true;
  if ADODataSet_Statistics.RecordCount>0 then
     Memo_statistics.Lines.Add('Выдано документов, всего: '+ADODataSet_Statistics.FieldByName('R').AsString);
  ADODataSet_Statistics.Active:=false;
  Memo_statistics.Lines.Add('');
  ADODataSet_Statistics.CommandText:='SELECT Count(*) AS R FROM EXPERTIZE_REESTR;';
  ADODataSet_Statistics.Active:=true;
  if ADODataSet_Statistics.RecordCount>0 then
     Memo_statistics.Lines.Add('Количество записей по экспертизам, всего: '+ADODataSet_Statistics.FieldByName('R').AsString);
  Memo_statistics.Lines.Add('');
  Memo_statistics.Lines.Add('');
  //год
  DecodeDate(StartOfTheYear(Date()), LYear, LMonth, LDay);
  Db:=IntTostr(LMonth)+'/'+IntTostr(LDay)+'/'+IntTostr(LYear);
  DecodeDate(EndOfTheYear(Date()), LYear, LMonth, LDay);
  De:=IntTostr(LMonth)+'/'+IntTostr(LDay)+'/'+IntTostr(LYear);
  ADODataSet_Statistics.Active:=false;
  ADODataSet_Statistics.CommandText:='SELECT Sum(KOLOBR) AS R FROM REGISTRATION WHERE ((DATEPROTOKOL>=#'+Db+'#) AND (DATEPROTOKOL<=#'+De+'#))';
  ADODataSet_Statistics.Active:=true;
  if ADODataSet_Statistics.RecordCount>0 then
     Memo_statistics.Lines.Add('Зарегистрировано образцов за текущий год: '+ADODataSet_Statistics.FieldByName('R').AsString);
  ADODataSet_Statistics.Active:=false;
  Memo_statistics.Lines.Add('');
  ADODataSet_Statistics.Active:=false;
  ADODataSet_Statistics.CommandText:='SELECT Count(*) AS R FROM REGISTRATION WHERE ((DATEPROTOKOL>=#'+Db+'#) AND (DATEPROTOKOL<=#'+De+'#))';
  ADODataSet_Statistics.Active:=true;
  if ADODataSet_Statistics.RecordCount>0 then
     Memo_statistics.Lines.Add('Выдано документов за текущий год: '+ADODataSet_Statistics.FieldByName('R').AsString);
  ADODataSet_Statistics.Active:=false;
  Memo_statistics.Lines.Add('');
  ADODataSet_Statistics.CommandText:='SELECT Count(*) AS R FROM EXPERTIZE_REESTR WHERE ((DATEEXPERTIZE>=#'+Db+'#) AND (DATEEXPERTIZE<=#'+De+'#))';
  ADODataSet_Statistics.Active:=true;
  if ADODataSet_Statistics.RecordCount>0 then
     Memo_statistics.Lines.Add('Количество записей по экспертизам за текущий год: '+ADODataSet_Statistics.FieldByName('R').AsString);
  ADODataSet_Statistics.Active:=false;
  Memo_statistics.Lines.Add('');
  Memo_statistics.Lines.Add('');
  //месяц
  DecodeDate(StartOfTheMonth(Date()), LYear, LMonth, LDay);
  Db:=IntTostr(LMonth)+'/'+IntTostr(LDay)+'/'+IntTostr(LYear);
  DecodeDate(EndOfTheMonth(Date()), LYear, LMonth, LDay);
  De:=IntTostr(LMonth)+'/'+IntTostr(LDay)+'/'+IntTostr(LYear);
  ADODataSet_Statistics.Active:=false;
  ADODataSet_Statistics.CommandText:='SELECT Sum(KOLOBR) AS R FROM REGISTRATION WHERE ((DATEPROTOKOL>=#'+Db+'#) AND (DATEPROTOKOL<=#'+De+'#))';
  ADODataSet_Statistics.Active:=true;
  if ADODataSet_Statistics.RecordCount>0 then
     Memo_statistics.Lines.Add('Зарегистрировано образцов за текущий месяц: '+ADODataSet_Statistics.FieldByName('R').AsString);
  ADODataSet_Statistics.Active:=false;
  Memo_statistics.Lines.Add('');
  ADODataSet_Statistics.Active:=false;
  ADODataSet_Statistics.CommandText:='SELECT Count(*) AS R FROM REGISTRATION WHERE ((DATEPROTOKOL>=#'+Db+'#) AND (DATEPROTOKOL<=#'+De+'#))';
  ADODataSet_Statistics.Active:=true;
  if ADODataSet_Statistics.RecordCount>0 then
     Memo_statistics.Lines.Add('Выдано документов за текущий месяц: '+ADODataSet_Statistics.FieldByName('R').AsString);
  ADODataSet_Statistics.Active:=false;
  Memo_statistics.Lines.Add('');
  ADODataSet_Statistics.CommandText:='SELECT Count(*) AS R FROM EXPERTIZE_REESTR WHERE ((DATEEXPERTIZE>=#'+Db+'#) AND (DATEEXPERTIZE<=#'+De+'#))';
  ADODataSet_Statistics.Active:=true;
  if ADODataSet_Statistics.RecordCount>0 then
     Memo_statistics.Lines.Add('Количество записей по экспертизам за текущий месяц: '+ADODataSet_Statistics.FieldByName('R').AsString);
  ADODataSet_Statistics.Active:=false;
  Memo_statistics.Lines.Add('');
  Timer2.Enabled:=false;
  WiiProgressBar1.Visible:=false;
end;

procedure TFrmMainWindow.WiiProgressBar1Click(Sender: TObject);
begin
   Timer2.Enabled:=true;
end;

end.
