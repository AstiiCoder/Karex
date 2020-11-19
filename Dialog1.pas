unit Dialog1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxLookAndFeelPainters, AdvGlowButton, cxCheckBox, cxGroupBox,
  cxRadioGroup, cxTextEdit, StdCtrls, cxButtons, cxDropDownEdit, cxCalc,
  cxMaskEdit, cxCalendar, cxControls, cxContainer, cxEdit, cxLabel,
  ComCtrls, ExtCtrls, AdvPanel, DB, ADODB, cxLookupEdit, cxDBLookupEdit,
  cxDBLookupComboBox, ComObj, Menus, cxGraphics, cxImageComboBox, cxMemo;

type
  TFrmDialog1 = class(TForm)
    DialogWizard: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    TabSheet5: TTabSheet;
    TabSheet6: TTabSheet;
    TabSheet7: TTabSheet;
    TabSheet8: TTabSheet;
    TabSheet9: TTabSheet;
    TabSheet10: TTabSheet;
    TabSheet11: TTabSheet;
    LabelNum: TcxLabel;
    LabelDlgName: TcxLabel;
    LabelDate: TcxLabel;
    DateEdit1: TcxDateEdit;
    CalcEdit1: TcxCalcEdit;
    cxLabel1: TcxLabel;
    LabelNameUpr: TcxLabel;
    ButtonNameUpr: TcxButton;
    LabelNumEtiketki: TcxLabel;
    LabelDate3: TcxLabel;
    DateEdit3: TcxDateEdit;
    LabelPunktName: TcxLabel;
    ButtonPunktName: TcxButton;
    LabelSpetsialist: TcxLabel;
    ButtonBoxSpetsialist: TcxButton;
    TextEditOrgName: TcxTextEdit;
    TextEditAdres: TcxTextEdit;
    cxLabel4: TcxLabel;
    LabelNameClient: TcxLabel;
    cxLabel5: TcxLabel;
    TextEditDogovorNum: TcxTextEdit;
    cxButton1: TcxButton;
    LabelOrgName: TcxLabel;
    LabelAdres: TcxLabel;
    cxLabel6: TcxLabel;
    LabelZayavkaNum: TcxLabel;
    LabelDateEdit4: TcxLabel;
    DateEdit4: TcxDateEdit;
    cxLabel9: TcxLabel;
    ButtonSprPartii: TcxButton;
    cxLabel10: TcxLabel;
    ButtonSprKlass: TcxButton;
    cxLabel12: TcxLabel;
    ButtonSprProishozd: TcxButton;
    cxLabel13: TcxLabel;
    CalcEditEdIzm: TcxCalcEdit;
    ButtonSprVes: TcxButton;
    ButtonSprEdIzm: TcxButton;
    cxLabel14: TcxLabel;
    cxLabel15: TcxLabel;
    ButtonSprObrazets: TcxButton;
    cxLabel16: TcxLabel;
    CalcEditVes: TcxCalcEdit;
    cxLabel17: TcxLabel;
    cxLabel18: TcxLabel;
    cxLabel19: TcxLabel;
    LabelKolich: TcxLabel;
    CalcEditKolich: TcxCalcEdit;
    CalcEditShtuk: TcxCalcEdit;
    cxLabel20: TcxLabel;
    cxLabel21: TcxLabel;
    ButtonSprTransport: TcxButton;
    cxLabel22: TcxLabel;
    TextEditNazvTransp: TcxTextEdit;
    cxLabel23: TcxLabel;
    RadioGroupPunkt: TcxRadioGroup;
    TextEditMesto: TcxTextEdit;
    cxLabel24: TcxLabel;
    cxCheckBox1: TcxCheckBox;
    cxCheckBox2: TcxCheckBox;
    cxCheckBox3: TcxCheckBox;
    cxLabel25: TcxLabel;
    cxCheckBox4: TcxCheckBox;
    cxCheckBox5: TcxCheckBox;
    LabelInfo: TcxLabel;
    cxLabel26: TcxLabel;
    ButtonForward: TAdvGlowButton;
    ButtonCancel: TAdvGlowButton;
    ButtonBackward: TAdvGlowButton;
    AdvPanel1: TAdvPanel;
    RadioGroupImport: TcxRadioGroup;
    TextEditSort: TcxTextEdit;
    CheckBoxSort: TcxCheckBox;
    LabelPodskazka: TcxLabel;
    DataSource1: TDataSource;
    ADODataSet1: TADODataSet;
    ComboBoxSpetsialist: TcxLookupComboBox;
    TextEditNumEtiketki: TcxTextEdit;
    ComboBoxPunkt: TcxLookupComboBox;
    DataSource2: TDataSource;
    ADODataSet2: TADODataSet;
    ADODataSet3: TADODataSet;
    DataSource3: TDataSource;
    ComboBoxUprNadz: TcxLookupComboBox;
    ADODataSet4: TADODataSet;
    DataSource4: TDataSource;
    ComboBoxOrganiz: TcxLookupComboBox;
    ADODataSet5: TADODataSet;
    DataSource5: TDataSource;
    TextEditZayavkaNum: TcxTextEdit;
    ComboBoxObrazets: TcxLookupComboBox;
    ADODataSet6: TADODataSet;
    DataSource6: TDataSource;
    ComboBoxPartii: TcxLookupComboBox;
    ComboBoxKlass: TcxLookupComboBox;
    ADODataSet7: TADODataSet;
    DataSource7: TDataSource;
    ComboBoxProishozhd: TcxLookupComboBox;
    ADODataSet8: TADODataSet;
    DataSource8: TDataSource;
    ComboBoxVes: TcxLookupComboBox;
    ComboBoxEdIzm: TcxLookupComboBox;
    ADODataSet9: TADODataSet;
    DataSource9: TDataSource;
    ADODataSet10: TADODataSet;
    DataSource10: TDataSource;
    ComboBoxTransport: TcxLookupComboBox;
    ADODataSet11: TADODataSet;
    DataSource11: TDataSource;
    ADOCommand1: TADOCommand;
    ADODataSet12: TADODataSet;
    CheckBoxNetEtiketki: TcxCheckBox;
    CheckBoxNetZayavki: TcxCheckBox;
    CheckBoxShtuk: TcxCheckBox;
    TextEditAbbr: TcxTextEdit;
    ADODataSet13: TADODataSet;
    DataSource13: TDataSource;
    LabelNazvDok: TcxLabel;
    ComboBoxNazvDok: TcxLookupComboBox;
    ButtonNazvDok: TcxButton;
    cxCheckBox6: TcxCheckBox;
    TabSheet12: TTabSheet;
    cxLabel2: TcxLabel;
    cxLabel3: TcxLabel;
    ComboBoxObrazetsOtech: TcxLookupComboBox;
    cxButton2: TcxButton;
    cxLabel7: TcxLabel;
    cxLabel8: TcxLabel;
    ComboBoxKlassOtech: TcxLookupComboBox;
    cxButton4: TcxButton;
    cxLabel11: TcxLabel;
    cxButton5: TcxButton;
    cxLabel27: TcxLabel;
    CalcEditVesOtech: TcxCalcEdit;
    ComboBoxVesOtech: TcxLookupComboBox;
    cxButton6: TcxButton;
    LabelKolvoMest: TcxLabel;
    CalcEditEdIzmOtech: TcxCalcEdit;
    ComboBoxEdIzmOtech: TcxLookupComboBox;
    cxButton7: TcxButton;
    RadioGroupPartyPole: TcxRadioGroup;
    TextEditProishOtech: TcxTextEdit;
    TextEditKodNasPunkt: TcxTextEdit;
    MemoSort: TcxMemo;
    cxLabel29: TcxLabel;
    ButtonMemoSize: TcxButton;
    ComboBoxProdukts: TcxLookupComboBox;
    cxButton3: TcxButton;
    ButtonExcelTool: TAdvGlowButton;
    CheckBoxExpToRF: TcxCheckBox;
    CheckBoxExpToExp: TcxCheckBox;
    LabelRegNazn: TcxLabel;
    RadioGroupNumType: TcxRadioGroup;
    procedure ButtonCancelClick(Sender: TObject);
    procedure ButtonForwardClick(Sender: TObject);
    procedure ButtonBackwardClick(Sender: TObject);
    procedure CheckBoxSortPropertiesChange(Sender: TObject);
    procedure cxRadioGroup1PropertiesChange(Sender: TObject);
    procedure ButtonSprVesClick(Sender: TObject);
    procedure ButtonSprEdIzmClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ButtonNameUprClick(Sender: TObject);
    procedure ButtonNameUprExit(Sender: TObject);
    procedure ButtonPunktNameClick(Sender: TObject);
    procedure ButtonPunktNameExit(Sender: TObject);
    procedure ButtonBoxSpetsialistClick(Sender: TObject);
    procedure ButtonBoxSpetsialistExit(Sender: TObject);
    procedure cxButton1Click(Sender: TObject);
    procedure cxButton1Exit(Sender: TObject);
    procedure ButtonSprObrazetsClick(Sender: TObject);
    procedure ButtonSprObrazetsExit(Sender: TObject);
    procedure ButtonSprPartiiClick(Sender: TObject);
    procedure ButtonSprPartiiExit(Sender: TObject);
    procedure ButtonSprKlassClick(Sender: TObject);
    procedure ButtonSprKlassExit(Sender: TObject);
    procedure ButtonSprVesExit(Sender: TObject);
    procedure ButtonSprEdIzmExit(Sender: TObject);
    procedure ButtonSprTransportClick(Sender: TObject);
    procedure ButtonSprTransportExit(Sender: TObject);
    procedure ComboBoxOrganizPropertiesChange(Sender: TObject);
    procedure ComboBoxOrganizPropertiesEditValueChanged(Sender: TObject);
    procedure DateEdit1PropertiesEditValueChanged(Sender: TObject);
    procedure RadioGroupImportPropertiesEditValueChanged(Sender: TObject);
    procedure CheckBoxNetEtiketkiPropertiesEditValueChanged(
      Sender: TObject);
    procedure CheckBoxNetZayavkiPropertiesEditValueChanged(
      Sender: TObject);
    procedure CheckBoxShtukPropertiesEditValueChanged(Sender: TObject);
    procedure ButtonNazvDokClick(Sender: TObject);
    procedure ButtonNazvDokExit(Sender: TObject);
    procedure ComboBoxNazvDokPropertiesEditValueChanged(Sender: TObject);
    procedure ComboBoxUprNadzPropertiesEditValueChanged(Sender: TObject);
    procedure cxButton5Click(Sender: TObject);
    procedure ButtonMemoSizeClick(Sender: TObject);
    procedure RadioGroupPartyPolePropertiesEditValueChanged(
      Sender: TObject);
    procedure TextEditSortPropertiesEditValueChanged(Sender: TObject);
    procedure TextEditNazvTranspPropertiesEditValueChanged(
      Sender: TObject);
    procedure TextEditMestoPropertiesEditValueChanged(Sender: TObject);
    procedure MemoSortPropertiesEditValueChanged(Sender: TObject);
    procedure TextEditProishOtechPropertiesEditValueChanged(
      Sender: TObject);
    procedure ButtonSprProishozdClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ButtonExcelToolClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ButtonSprProishozdExit(Sender: TObject);
    procedure TextEditNumEtiketkiPropertiesEditValueChanged(
      Sender: TObject);
    procedure DateEdit3PropertiesEditValueChanged(Sender: TObject);
    procedure TextEditDogovorNumPropertiesEditValueChanged(
      Sender: TObject);
    procedure TextEditZayavkaNumPropertiesEditValueChanged(
      Sender: TObject);
    procedure DateEdit4PropertiesEditValueChanged(Sender: TObject);
    procedure CalcEditVesPropertiesEditValueChanged(Sender: TObject);
    procedure CalcEditEdIzmPropertiesEditValueChanged(Sender: TObject);
    procedure CalcEditKolichPropertiesEditValueChanged(Sender: TObject);
    procedure CalcEditShtukPropertiesEditValueChanged(Sender: TObject);
    procedure cxCalcEdit1PropertiesChange(Sender: TObject);
    procedure cxCalcEdit2PropertiesEditValueChanged(Sender: TObject);
    procedure RadioGroupNumTypePropertiesChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmDialog1: TFrmDialog1;
  TmpObj: TcxLookupComboBox;

implementation

uses
  Grid1, MainWindow, ProishObrOtech, cxGridTableView, ToolExcel;

{$R *.dfm}

function StrReplace(const Str, Str1, Str2: string): string;
// str - исходная строка
// str1 - подстрока, подлежащая замене
// str2 - заменяющая строка
var
  P, L: Integer;
begin
  Result := str;
  L := Length(Str1);
  repeat
    P := Pos(Str1, Result); // ищем подстроку
    if P > 0 then
    begin
      Delete(Result, P, L); // удаляем ее
      Insert(Str2, Result, P); // вставляем новую
    end;
  until P = 0;
end;

function SQL_Float_String(Value: double): string;
//адаптация чисел к разделителям дробной части
var
  OldSeparator: Char;
begin
  OldSeparator := DecimalSeparator;
  DecimalSeparator := '.';
  Result := FloatToStr(Value);
  DecimalSeparator := OldSeparator;
end;

procedure TFrmDialog1.ButtonCancelClick(Sender: TObject);
begin
  close;
end;

procedure TFrmDialog1.ButtonForwardClick(Sender: TObject);
var
  SQLstr, a, b, c, DayProtocol: string;
  LYear, LMonth, LDay: Word;
  Excel : variant;
begin
  sleep(300);
  if DialogWizard.ActivePage = TabSheet1 then
    begin
      DialogWizard.ActivePage := TabSheet2;
      ButtonBackward.Enabled:=true;
      LastNumProtForManual:=CalcEdit1.EditValue;
      LabelPodskazka.Caption:='Шаг 2';
      TabSheet2.Show;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet2 then
    begin
      DialogWizard.ActivePage := TabSheet3;
      LabelPodskazka.Caption:='Шаг 3';
      TabSheet3.Show;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet3 then
    begin
      DialogWizard.ActivePage := TabSheet4;
      LabelPodskazka.Caption:='Шаг 4';
      TabSheet4.Show;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet4 then
    begin
      if RadioGroupImport.ItemIndex=0 then
        begin
          DialogWizard.ActivePage := TabSheet5;
          LabelPodskazka.Caption:='Шаг 5';
          TabSheet5.Show;
          exit;
        end
      else
        begin
          DialogWizard.ActivePage := TabSheet12;
          LabelPodskazka.Caption:='Шаг 5';
          TabSheet12.Show;
          exit;
        end
    end;
  if DialogWizard.ActivePage = TabSheet5 then
    begin
      DialogWizard.ActivePage := TabSheet6;
      LabelPodskazka.Caption:='Шаг 6';
      TabSheet6.Show;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet6 then
    begin
      DialogWizard.ActivePage := TabSheet7;
      LabelPodskazka.Caption:='Шаг 7';
      TabSheet7.Show;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet7 then
    begin
      DialogWizard.ActivePage := TabSheet8;
      LabelPodskazka.Caption:='Шаг 8';
      if RadioGroupImport.ItemIndex=0 then
        begin //импорт
          LabelRegNazn.Visible:=false;
          CheckBoxExpToRF.Visible:=false;
          CheckBoxExpToExp.Visible:=false;
        end
      else
        begin //отечественные
          LabelRegNazn.Visible:=true;
          CheckBoxExpToRF.Visible:=true;
          CheckBoxExpToExp.Visible:=true;
        end;
      TabSheet8.Show;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet8 then
    begin
      DialogWizard.ActivePage := TabSheet9;
      LabelPodskazka.Caption:='Шаг 9';
      TabSheet9.Show;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet9 then
    begin
      DialogWizard.ActivePage := TabSheet10;
      LabelPodskazka.Caption:='Шаг 10';
      TabSheet10.Show;
      ButtonForward.Enabled:=true;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet10 then
    begin
      if TextEditMesto.Text='' then TextEditMesto.Text:='нет';
      //проверка, что такой протокол уже занесён, этой датой и этого происхождения
      FrmMainWindow.ADODataSetLockRegistration.Active:=false;
      DecodeDate(DateEdit3.Date, LYear, LMonth, LDay);
      DayProtocol:=IntTostr(LMonth)+'/'+IntTostr(LDay)+'/'+IntTostr(LYear);
      SQLstr:='SELECT REGISTRATION.DATEPROTOKOL, REGISTRATION.NUMPROTOKOL FROM REGISTRATION WHERE (((REGISTRATION.DATEPROTOKOL)="'+DayProtocol+'") AND ((REGISTRATION.NUMPROTOKOL)='+IntToStr(CalcEdit1.EditValue)+') AND ((REGISTRATION.ISIMPORT)='+IntToStr(RadioGroupImport.ItemIndex)+'))';
      FrmMainWindow.ADODataSetLockRegistration.CommandText:=SQLstr;
        try
          FrmMainWindow.ADODataSetLockRegistration.Active:=true;
          if FrmMainWindow.ADODataSetLockRegistration.RecordCount>0 then
            begin
                showmessage('Протокол №'+IntToStr(CalcEdit1.EditValue)+' занесён в базу!');
                close;
            end;
        except
        end;
      FrmMainWindow.ADODataSetLockRegistration.Active:=false;
      if RadioGroupImport.ItemIndex=0 then
        begin //импорт
      SQLstr:='INSERT INTO REGISTRATION (ISIMPORT, NUMPROTOKOL, DATEPROTOKOL, DOСTYPE, UPRSHNADZ, PUNKTKARRAST, SPETSIALIST, NUMETIKETKI, ETIKETKADATE, ORGANIZATION, NUMDOGOVOR, NUMZAYAVKA, DATEZAYAVKA, NAMEOBRAZTSA, NAMEPATRY, KLASSIFIKATION, SORT, PROISHOZHD,';
      SQLstr:=SQLstr+' VES, EDVES, KOL, EDIZM, KOLOBR, SHTUKVOBR, VIDTRANSPORT, NAZVTRANSPORT, PUNKTNAZNACH, VUPRSHNADZ, VCLIENT, VARCHIV, PUPRSHNADZ, PCLIENT, PREFCENTR, PUNKTNNO)';
      SQLstr:=SQLstr+' SELECT '+IntToStr(RadioGroupImport.ItemIndex)+' AS B1, '+IntToStr(CalcEdit1.EditValue)+' AS B2, "'+DateToStr(DateEdit1.Date)+'" AS B3, '+IntToStr(ComboBoxNazvDok.EditValue)+' AS B3_1, '+IntToStr(ComboBoxUprNadz.EditValue)+' AS B4, '+IntToStr(ComboBoxPunkt.EditValue)+' AS B5,';
      SQLstr:=SQLstr+' '+IntToStr(ComboBoxSpetsialist.EditValue)+' AS B6, "'+TextEditNumEtiketki.Text+'" AS B7, "'+DateToStr(DateEdit3.Date)+'" AS B8, '+IntToStr(ComboBoxOrganiz.EditValue)+' AS B9,';
      SQLstr:=SQLstr+' "'+TextEditDogovorNum.Text+'" AS B10, "'+TextEditZayavkaNum.Text+'" AS B11, "'+DateToStr(DateEdit4.Date)+'" AS B12, '+IntToStr(ComboBoxObrazets.EditValue)+' AS B13,';
      SQLstr:=SQLstr+' '+IntToStr(ComboBoxPartii.EditValue)+' AS B14, '+IntToStr(ComboBoxKlass.EditValue)+' AS B15, "'+TextEditSort.Text+'" AS B16, '+IntToStr(ComboBoxProishozhd.EditValue)+' AS B17,';
      SQLstr:=SQLstr+' '+SQL_Float_String(CalcEditVes.EditValue)+' AS B18, '+IntToStr(ComboBoxVes.EditValue)+' AS B19, '+IntToStr(CalcEditEdIzm.EditValue)+' AS B20, '+IntToStr(ComboBoxEdIzm.EditValue)+' AS B21,';
      SQLstr:=SQLstr+' '+IntToStr(CalcEditKolich.EditValue)+' AS B22, '+IntToStr(CalcEditShtuk.EditValue)+' AS B23, '+IntToStr(ComboBoxTransport.EditValue)+' AS B24, "'+StrReplace(TextEditNazvTransp.Text,'"','`')+'" AS B25,';
      SQLstr:=SQLstr+' "'+StrReplace(TextEditMesto.Text,'"','`')+'" AS B26, '+IntToStr(cxCheckBox1.EditValue)+' AS B27, '+IntToStr(cxCheckBox2.EditValue)+' AS B28, '+IntToStr(cxCheckBox3.EditValue)+' AS B29, '+IntToStr(cxCheckBox4.EditValue)+' AS B30,';
      SQLstr:=SQLstr+' '+IntToStr(cxCheckBox5.EditValue)+' AS B31, '+IntToStr(cxCheckBox6.EditValue)+' AS B32, '+IntToStr(RadioGroupPunkt.ItemIndex)+' AS B35';
      ADOCommand1.CommandText:=SQLstr;
          try
            ButtonForward.Enabled:=false;
            ADOCommand1.Execute;
          except
            showmessage('Не все обязательные поля заполнены! Проверьте правильность ввода данных.'+Chr(13)+Chr(13)+'Сделайте снимок экрана и покажите разработчикам для детального анализа ошибки.'+Chr(13)+SQLstr);
            exit;
          end;
        end
      else
        begin //отечественные
      SQLstr:='INSERT INTO REGISTRATION (ISIMPORT, NUMPROTOKOL, DATEPROTOKOL, DOСTYPE, UPRSHNADZ, PUNKTKARRAST, SPETSIALIST, NUMETIKETKI, ETIKETKADATE, ORGANIZATION, NUMDOGOVOR, NUMZAYAVKA, DATEZAYAVKA, NAMEOBRAZTSA, NAMEPATRY, KLASSIFIKATION, SORT, PROISHOZHD,';
      SQLstr:=SQLstr+' VES, EDVES, KOL, EDIZM, KOLOBR, SHTUKVOBR, VIDTRANSPORT, NAZVTRANSPORT, PUNKTNAZNACH, VUPRSHNADZ, VCLIENT, VARCHIV, PUPRSHNADZ, PCLIENT, PREFCENTR, NAZNRF, NAZNEXP, PUNKTNNO)';
      SQLstr:=SQLstr+' SELECT '+IntToStr(RadioGroupImport.ItemIndex)+' AS B1, '+IntToStr(CalcEdit1.EditValue)+' AS B2, "'+DateToStr(DateEdit1.Date)+'" AS B3, '+IntToStr(ComboBoxNazvDok.EditValue)+' AS B3_1, '+IntToStr(ComboBoxUprNadz.EditValue)+' AS B4, '+IntToStr(ComboBoxPunkt.EditValue)+' AS B5,';
      SQLstr:=SQLstr+' '+IntToStr(ComboBoxSpetsialist.EditValue)+' AS B6, "'+TextEditNumEtiketki.Text+'" AS B7, "'+DateToStr(DateEdit3.Date)+'" AS B8, '+IntToStr(ComboBoxOrganiz.EditValue)+' AS B9,';
      SQLstr:=SQLstr+' "'+TextEditDogovorNum.Text+'" AS B10, "'+TextEditZayavkaNum.Text+'" AS B11, "'+DateToStr(DateEdit4.Date)+'" AS B12, '+IntToStr(ComboBoxObrazetsOtech.EditValue)+' AS B13,';
      SQLstr:=SQLstr+' '+IntToStr(ComboBoxProdukts.EditValue)+' AS B14, '+IntToStr(ComboBoxKlassOtech.EditValue)+' AS B15, "'+TextEditSort.Text+'" AS B16, '+TextEditKodNasPunkt.Text+' AS B17,';
      SQLstr:=SQLstr+' '+SQL_Float_String(CalcEditVesOtech.EditValue)+' AS B18, '+IntToStr(ComboBoxVesOtech.EditValue)+' AS B19, '+IntToStr(CalcEditEdIzmOtech.EditValue)+' AS B20, '+IntToStr(ComboBoxEdIzmOtech.EditValue)+' AS B21,';
      SQLstr:=SQLstr+' '+IntToStr(CalcEditKolich.EditValue)+' AS B22, '+IntToStr(CalcEditShtuk.EditValue)+' AS B23, '+IntToStr(ComboBoxTransport.EditValue)+' AS B24, "'+StrReplace(TextEditNazvTransp.Text,'"','`')+'" AS B25,';
      SQLstr:=SQLstr+' "'+StrReplace(TextEditMesto.Text,'"','`')+'" AS B26, '+IntToStr(cxCheckBox1.EditValue)+' AS B27, '+IntToStr(cxCheckBox2.EditValue)+' AS B28, '+IntToStr(cxCheckBox3.EditValue)+' AS B29, '+IntToStr(cxCheckBox4.EditValue)+' AS B30,';
      SQLstr:=SQLstr+' '+IntToStr(cxCheckBox5.EditValue)+' AS B31, '+IntToStr(cxCheckBox6.EditValue)+' AS B32, '+IntToStr(CheckBoxExpToRF.EditValue)+' AS B33, '+IntToStr(CheckBoxExpToExp.EditValue)+' AS B34, '+IntToStr(RadioGroupPunkt.ItemIndex*(-1))+' AS B35';
      ADOCommand1.CommandText:=SQLstr;
          try
            ButtonForward.Enabled:=false;
            ADOCommand1.Execute;
          except
            showmessage('Не все обязательные поля заполнены! Проверьте правильность ввода данных.');
            exit;
          end;
      //определяем последний номер ID регистрации
      ADODataSet12.Active:=false;
      ADODataSet12.CommandText:='SELECT Max(IDREGISTRATION) AS LASTID FROM REGISTRATION WHERE (((ISIMPORT)=1))';
      ADODataSet12.Active:=true;
      c:=IntToStr(ADODataSet12.FieldByName('LASTID').AsInteger+1);
      ADODataSet12.Active:=false;
      MemoSort.Lines.Delimiter:='|';
      a:=MemoSort.Text;
      b:=TextEditProishOtech.Text;
      if length(a)=0 then a:='нет';
      if length(b)=0 then b:='нет';
      ADOCommand1.CommandText:='INSERT INTO OTECHMEMOSORT (SORTTMPROCH, PROISH, IDREGISTRATION) SELECT "'+a+'" AS SO, "'+b+'" AS PR, "'+c+'" AS IDA';
      ADOCommand1.Execute;
        end;
      DialogWizard.ActivePage := TabSheet11;
      LabelPodskazka.Caption:='Шаг 11';
      TabSheet11.Show;
      ButtonBackward.Enabled:=false;
      ButtonForward.Caption:='Готово!';
      ButtonForward.Enabled:=true;
      ButtonCancel.Enabled:=false;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet11 then
    begin
       if Oeape=true then
           begin // выгрузка протокола в EXCEL
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
            ExcelApp.Application.DisplayAlerts := false; // спрашивать в Excel сохранять ли файл после изменений
            ExcelApp.Visible:=true;
            Screen.Cursor:=crDefault;
            Workbook := ExcelApp.WorkBooks.Open(ExtractFilePath(Application.ExeName)+'Templates\Протокол.xls',0,true);
            if RadioGroupImport.ItemIndex=1 then
              ExcelApp.WorkBooks[1].WorkSheets[1].Cells[3, 8] :='O-'+CalcEdit1.Text //номер протокола отечественного
            else
              ExcelApp.WorkBooks[1].WorkSheets[1].Cells[3, 8] :=CalcEdit1.Text; //номер протокола
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[5, 1] :=ComboBoxNazvDok.Text;  //название документа
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[12, 2] :=TextEditNumEtiketki.Text;  //номер документа
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[13, 2] :=DateToStr(DateEdit3.Date); //дата выдачи этикетки
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[14, 1] :=ComboBoxUprNadz.Text; //наим упр с/х надз
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[21, 1] :=ComboBoxPunkt.Text; //назв. пункта по карантиру растений
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[27, 1] :=ComboBoxSpetsialist.Text; //специалист
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[30, 2] :=TextEditZayavkaNum.Text; //заявка клиента
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[31, 2] :=DateToStr(DateEdit4.Date); //номер заявки
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[32, 1] :=TextEditAbbr.Text; //аббревиатура
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[33, 1] :=ComboBoxOrganiz.Text; //назв.организации
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[36, 1] :=TextEditAdres.Text; //адрес
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[5, 3] :=DateToStr(DateEdit1.Date); //дата оформил
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[7, 4] :=ComboBoxObrazets.Text; //наименование образца
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[19, 4] :=ComboBoxPartii.Text; //наименование партии
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[24, 4] :=TextEditSort.Text; //сорт
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[28, 4] :=CalcEditVes.Text+' '+ComboBoxVes.Text; //вес
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[30, 4] :=CalcEditEdIzm.Text+' '+ComboBoxEdIzm.Text; //мест
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[5, 5] :=CalcEditKolich.Value; //кол-во обр
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[7, 5] :=CalcEditShtuk.Value;  //штук в обр
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[5, 6] :=ComboBoxProishozhd.Text;  //происх обр
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[6, 6] :=ComboBoxTransport.Text;  //транспорт
            ExcelApp.WorkBooks[1].WorkSheets[1].Cells[7, 6] :=TextEditNazvTransp.Text;  //расшифровка
            if RadioGroupPunkt.ItemIndex=0 then
              begin
                ExcelApp.WorkBooks[1].WorkSheets[1].Cells[5, 7] :=TextEditAbbr.Text;  //аббревиатура
                ExcelApp.WorkBooks[1].WorkSheets[1].Cells[6, 7] :=ComboBoxOrganiz.Text;  //назв.организации
                ExcelApp.WorkBooks[1].WorkSheets[1].Cells[11, 7] :=TextEditAdres.Text;   //адрес организации
              end
            else
              begin
                ExcelApp.WorkBooks[1].WorkSheets[1].Cells[5, 7] :='';  //аббревиатура
                ExcelApp.WorkBooks[1].WorkSheets[1].Cells[6, 7] :=TextEditMesto.Text;  //назв.организации
                ExcelApp.WorkBooks[1].WorkSheets[1].Cells[11, 7] :='';   //адрес организации
              end;
            //защищаем лист от изменений
            ExcelApp.WorkBooks[1].WorkSheets[1].protect('344',true,true,false);
        end; // выгрузка протокола в EXCEL
      //возвращение на первый шаг
      DialogWizard.ActivePage := TabSheet1;
      LabelPodskazka.Caption:='Шаг 1 (следующая экспертиза)';
      ButtonBackward.Enabled:=false;
      ButtonForward.Caption:='Продолжить >';
      ButtonForward.Enabled:=true;
      ButtonCancel.Enabled:=true;
      //определяем дату следующего протокола
      if RadioGroupNumType.ItemIndex=0 then
         begin
          //если нумерация автоматическая
          ADODataSet12.Active:=false;
          //SQLstr:='SELECT Max(REGISTRATION.NUMPROTOKOL) AS NUM, Max(REGISTRATION.DATEPROTOKOL) AS DATA FROM REGISTRATION WHERE (((REGISTRATION.ISIMPORT)='+IntToStr(FrmDialog1.RadioGroupImport.ItemIndex)+'))';
          //SQLstr:=SQLstr+' HAVING (((Max(REGISTRATION.DATEPROTOKOL))>=#'+DateTekPerStart+'# And (Max(REGISTRATION.DATEPROTOKOL))<#'+DateTekPerFinish+'#)) ORDER BY Max(REGISTRATION.NUMPROTOKOL) DESC;';
          SQLstr:='SELECT REGISTRATION.NUMPROTOKOL AS NUM, REGISTRATION.DATEPROTOKOL AS DATA FROM REGISTRATION';
          SQLstr:=SQLstr+' WHERE (((REGISTRATION.ISIMPORT)='+IntToStr(FrmDialog1.RadioGroupImport.ItemIndex)+') AND ((REGISTRATION.DATEPROTOKOL)>=#'+DateTekPerStart+'# And (REGISTRATION.DATEPROTOKOL)<#'+DateTekPerFinish+'#)) ORDER BY REGISTRATION.NUMPROTOKOL DESC;';
          ADODataSet12.CommandText:=SQLstr;
          ADODataSet12.Active:=true;
          if ADODataSet12.RecordCount>0 then
            begin
              CalcEdit1.EditValue:=IntToStr(ADODataSet12.FieldByName('NUM').AsInteger+1);
              DateLastDog:=ADODataSet12.FieldByName('DATA').AsString;
            end;
         end
      else
          begin
            //если ручная нумерация протоколов
            LastNumProtForManual:=LastNumProtForManual+1;
            CalcEdit1.EditValue:=LastNumProtForManual;
          end;
      //запускаем Excel toolbar
      if Oeape=true then
         begin
            if FrmToolExcel= NIL then FrmToolExcel := TFrmToolExcel.Create(owner);
            FrmToolExcel.Show;
            ButtonExcelTool.Visible:=false;
         end;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet12 then
    begin
      DialogWizard.ActivePage := TabSheet6;
      LabelPodskazka.Caption:='Шаг 6';
      TabSheet6.Show;
      exit;
    end;
end;

procedure TFrmDialog1.ButtonBackwardClick(Sender: TObject);
begin
  if DialogWizard.ActivePage = TabSheet12 then
    begin
      DialogWizard.ActivePage := TabSheet4;
      LabelPodskazka.Caption:='Шаг 4';
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet10 then
    begin
      DialogWizard.ActivePage := TabSheet9;
      LabelPodskazka.Caption:='Шаг 9';
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet9 then
    begin
      DialogWizard.ActivePage := TabSheet8;
      LabelPodskazka.Caption:='Шаг 8';
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet8 then
    begin
      DialogWizard.ActivePage := TabSheet7;
      LabelPodskazka.Caption:='Шаг 7';
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet7 then
    begin
      DialogWizard.ActivePage := TabSheet6;
      LabelPodskazka.Caption:='Шаг 6';
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet6 then
    begin
      if RadioGroupImport.ItemIndex=0 then DialogWizard.ActivePage := TabSheet5 else DialogWizard.ActivePage := TabSheet12;
      LabelPodskazka.Caption:='Шаг 5';
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet5 then
    begin
      DialogWizard.ActivePage := TabSheet4;
      LabelPodskazka.Caption:='Шаг 4';
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet4 then
    begin
      DialogWizard.ActivePage := TabSheet3;
      LabelPodskazka.Caption:='Шаг 3';
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet3 then
    begin
      DialogWizard.ActivePage := TabSheet2;
      LabelPodskazka.Caption:='Шаг 2';
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet2 then
    begin
      DialogWizard.ActivePage := TabSheet1;
      LabelPodskazka.Caption:='Шаг 1';
      ButtonBackward.Enabled:=false;
      exit;
    end;
end;

procedure TFrmDialog1.CheckBoxSortPropertiesChange(Sender: TObject);
begin
  if CheckBoxSort.Checked=true then
    TextEditSort.Enabled:=true
  else
    begin
     TextEditSort.Enabled:=false;
     TextEditSort.Text:='нет';
    end;
end;

procedure TFrmDialog1.cxRadioGroup1PropertiesChange(Sender: TObject);
begin
  if RadioGroupPunkt.ItemIndex=1 then
    begin
      TextEditMesto.Enabled:=true;
      TextEditMesto.Visible:=true;
    end
  else
    begin
       TextEditMesto.Enabled:=false;
       TextEditMesto.Text:=ComboBoxOrganiz.SelText;
       TextEditMesto.Visible:=false;
    end;
end;

procedure TFrmDialog1.FormCreate(Sender: TObject);
begin
  //открываем датасеты
  ADODataSet1.Active:=true;
  ADODataSet2.Active:=true;
  ADODataSet3.Active:=true;
  ADODataSet4.Active:=true;
  ADODataSet5.Active:=true;
  ADODataSet6.Active:=true;
  ADODataSet7.Active:=true;
  ADODataSet8.Active:=true;
  ADODataSet9.Active:=true;
  ADODataSet10.Active:=true;
  ADODataSet11.Active:=true;
  ADODataSet13.Active:=true;
  //тип нумерации прококолов
  TypeNumeration:=0;
  RadioGroupNumType.Enabled:=MUReg;
  if MUReg=false then RadioGroupNumType.ItemIndex:=0;
end;

procedure TFrmDialog1.ButtonNameUprClick(Sender: TObject);
begin
  TblRezh:=1;
  TmpObj:=ComboBoxUprNadz;
  FrmMainWindow.ButtonShowUprNadzoraClick(sender);
end;

procedure TFrmDialog1.ButtonNameUprExit(Sender: TObject);
begin
  ADODataSet3.Active:=false;
  ADODataSet3.Active:=true;
end;

procedure TFrmDialog1.ButtonPunktNameClick(Sender: TObject);
begin
 TblRezh:=1;
 TmpObj:=ComboBoxPunkt;
 FrmMainWindow.ButtonShowPunktsClick(sender);
end;

procedure TFrmDialog1.ButtonPunktNameExit(Sender: TObject);
begin
  ADODataSet2.Active:=false;
  ADODataSet2.Active:=true;
end;

procedure TFrmDialog1.ButtonBoxSpetsialistClick(Sender: TObject);
begin
  TblRezh:=1;
  TmpObj:=ComboBoxSpetsialist;
  FrmMainWindow.ButtonShowSpetsialistClick(Sender);
end;

procedure TFrmDialog1.ButtonBoxSpetsialistExit(Sender: TObject);
begin
  ADODataSet1.Active:=false;
  ADODataSet1.Active:=true;
end;

procedure TFrmDialog1.cxButton1Click(Sender: TObject);
begin
  TblRezh:=1;
  TmpObj:=ComboBoxOrganiz;
  FrmMainWindow.ButtonShowClientsClick(Sender);
end;

procedure TFrmDialog1.cxButton1Exit(Sender: TObject);
begin
  ADODataSet4.Active:=false;
  ADODataSet4.Active:=true;
  ComboBoxOrganizPropertiesEditValueChanged(Sender);
  TextEditMesto.Text:='нет';
end;

procedure TFrmDialog1.ButtonSprObrazetsClick(Sender: TObject);
begin
  TblRezh:=1;
  TmpObj:=ComboBoxObrazets;
  FrmMainWindow.ButtonShowObrazetsClick(Sender);
end;

procedure TFrmDialog1.ButtonSprObrazetsExit(Sender: TObject);
begin
  ADODataSet5.Active:=false;
  ADODataSet5.Active:=true;
end;

procedure TFrmDialog1.ButtonSprPartiiClick(Sender: TObject);
begin
  TblRezh:=1;
  TmpObj:=ComboBoxPartii;
  FrmMainWindow.ButtonShowPartiiClick(Sender);
end;

procedure TFrmDialog1.ButtonSprPartiiExit(Sender: TObject);
begin
  ADODataSet6.Active:=false;
  ADODataSet6.Active:=true;
end;

procedure TFrmDialog1.ButtonSprKlassClick(Sender: TObject);
begin
  TblRezh:=1;
  TmpObj:=ComboBoxKlass;
  FrmMainWindow.ButtonShowKlassClick(Sender);
end;

procedure TFrmDialog1.ButtonSprKlassExit(Sender: TObject);
begin
  ADODataSet7.Active:=false;
  ADODataSet7.Active:=true;
end;

procedure TFrmDialog1.ButtonSprVesClick(Sender: TObject);
begin
  TblRezh:=1;
  TmpObj:=ComboBoxVes;
  FrmMainWindow.ButtonShowVesClick(Sender);
end;

procedure TFrmDialog1.ButtonSprVesExit(Sender: TObject);
begin
  ADODataSet9.Active:=false;
  ADODataSet9.Active:=true;
end;

procedure TFrmDialog1.ButtonSprEdIzmClick(Sender: TObject);
begin
  TblRezh:=1;
  TmpObj:=ComboBoxEdIzm;
  FrmMainWindow.ButtonShowEdIzmClick(Sender);
end;

procedure TFrmDialog1.ButtonSprEdIzmExit(Sender: TObject);
begin
  ADODataSet10.Active:=false;
  ADODataSet10.Active:=true;
end;

procedure TFrmDialog1.ButtonSprTransportClick(Sender: TObject);
begin
  TblRezh:=1;
  TmpObj:=ComboBoxTransport;
  FrmMainWindow.ButtonShowTransportClick(Sender);
end;

procedure TFrmDialog1.ButtonSprTransportExit(Sender: TObject);
begin
  ADODataSet11.Active:=false;
  ADODataSet11.Active:=true;
end;

procedure TFrmDialog1.ComboBoxOrganizPropertiesChange(Sender: TObject);
begin
  TextEditMesto.Text:=ComboBoxOrganiz.SelText;
end;

procedure TFrmDialog1.ComboBoxOrganizPropertiesEditValueChanged(
  Sender: TObject);
begin
  if ComboBoxOrganiz.Text='' then
    begin
      TextEditOrgName.Text:='';
      TextEditAdres.Text:='';
      TextEditAbbr.Text:='';
      TextEditDogovorNum.Text:='';
      exit;
    end;
  //определение реквизитов организации
  ADODataSet12.Active:=false;
  ADODataSet12.CommandText:='SELECT CLIENTS.*, ABBREVIATURA.ABBR, ABBREVIATURA.IDABBR FROM CLIENTS LEFT JOIN ABBREVIATURA ON CLIENTS.ABBREVCLIENT = ABBREVIATURA.IDABBR WHERE (((IDCLIENT)='+IntToStr(ComboBoxOrganiz.EditValue)+'));';
  ADODataSet12.Active:=true;
  if ADODataSet12.RecordCount>0 then
    begin
      ADODataSet12.First;
      TextEditOrgName.Text:=ADODataSet12.FieldByName('IDABBR').AsString;
      TextEditAdres.Text:=ADODataSet12.FieldByName('ADRESSCLIENT').AsString;
      TextEditAbbr.Text:=ADODataSet12.FieldByName('ABBR').AsString;
      try
        TextEditDogovorNum.Text:=ADODataSet12.FieldByName('DOGNUM').AsString;
      except
      end;
    end
  else
    begin
      TextEditOrgName.Text:='';
      TextEditAdres.Text:='';
      TextEditDogovorNum.Text:='';
    end
end;

procedure TFrmDialog1.DateEdit1PropertiesEditValueChanged(Sender: TObject);
begin
  try
    if DateEdit1.Date<StrToDate(DateLastDog) then
      begin
        showmessage('Нельзя создать протокол за указанную дата, поскольку предыдущий №'+IntToStr(CalcEdit1.EditValue-1)+' был закреплён за договором от '+DateLastDog);
        DateEdit1.Date:=Date;
      end;
   except
   end;
end;

procedure TFrmDialog1.RadioGroupImportPropertiesEditValueChanged(
  Sender: TObject);
var
 SQLstr: string;
begin
 {if RadioGroupImport.ItemIndex=1 then
  begin
    showmessage('Данная версия программы поддерживет только "импортный"');
    RadioGroupImport.ItemIndex:=0;
    exit;
  end;}
 ADODataSet12.Active:=false;
 //SQLstr:='SELECT Max(REGISTRATION.NUMPROTOKOL) AS NUM, Max(REGISTRATION.DATEPROTOKOL) AS DATA FROM REGISTRATION WHERE (((REGISTRATION.ISIMPORT)='+IntToStr(FrmDialog1.RadioGroupImport.ItemIndex)+'))';
 //SQLstr:=SQLstr+' HAVING (((Max(REGISTRATION.DATEPROTOKOL))>=#'+DateTekPerStart+'# And (Max(REGISTRATION.DATEPROTOKOL))<#'+DateTekPerFinish+'#)) ORDER BY Max(REGISTRATION.NUMPROTOKOL) DESC;';
 SQLstr:='SELECT REGISTRATION.NUMPROTOKOL AS NUM, REGISTRATION.DATEPROTOKOL AS DATA FROM REGISTRATION';
 SQLstr:=SQLstr+' WHERE (((REGISTRATION.ISIMPORT)='+IntToStr(FrmDialog1.RadioGroupImport.ItemIndex)+') AND ((REGISTRATION.DATEPROTOKOL)>=#'+DateTekPerStart+'# And (REGISTRATION.DATEPROTOKOL)<#'+DateTekPerFinish+'#)) ORDER BY REGISTRATION.NUMPROTOKOL DESC;';
 ADODataSet12.CommandText:=SQLstr;
 ADODataSet12.Active:=true;
 CalcEdit1.EditValue:=IntToStr(ADODataSet12.FieldByName('NUM').AsInteger+1);
 DateLastDog:=ADODataSet12.FieldByName('DATA').AsString;
 DateEdit1.Date:=Date;
end;

procedure TFrmDialog1.CheckBoxNetEtiketkiPropertiesEditValueChanged(
  Sender: TObject);
begin
  if CheckBoxNetEtiketki.Checked=true then
    begin
      ComboBoxNazvDok.Enabled:=false;
      ComboBoxUprNadz.Enabled:=false;
      ComboBoxPunkt.Enabled:=false;
      ComboBoxSpetsialist.Enabled:=false;
      TextEditNumEtiketki.Enabled:=false;
      DateEdit3.Enabled:=false;
      //значения в пустые
      ComboBoxNazvDok.EditValue:=0;
      ComboBoxUprNadz.EditValue:=0;
      ComboBoxPunkt.EditValue:=0;
      ComboBoxSpetsialist.EditValue:=0;
      TextEditNumEtiketki.EditValue:=0;
      DateEdit3.Date:=Date;
      //подписи тоже не активны
      LabelNazvDok.Enabled:=false;
      LabelNameUpr.Enabled:=false;
      LabelPunktName.Enabled:=false;
      LabelSpetsialist.Enabled:=false;
      LabelNumEtiketki.Enabled:=false;
      LabelDate3.Enabled:=false;
      //кнопки
      ButtonNazvDok.Enabled:=false;
      ButtonNameUpr.Enabled:=false;
      ButtonPunktName.Enabled:=false;
      ButtonBoxSpetsialist.Enabled:=false;
    end
  else
    begin
      ComboBoxNazvDok.Enabled:=true;
      ComboBoxUprNadz.Enabled:=true;
      ComboBoxPunkt.Enabled:=true;
      ComboBoxSpetsialist.Enabled:=true;
      TextEditNumEtiketki.Enabled:=true;
      DateEdit3.Enabled:=true;
      LabelNazvDok.Enabled:=true;
      LabelNameUpr.Enabled:=true;
      LabelPunktName.Enabled:=true;
      LabelSpetsialist.Enabled:=true;
      LabelNumEtiketki.Enabled:=true;
      LabelDate3.Enabled:=true;
      //кнопки
      ButtonNazvDok.Enabled:=true;
      ButtonNameUpr.Enabled:=true;
      ButtonPunktName.Enabled:=true;
      ButtonBoxSpetsialist.Enabled:=true;
    end;
end;

procedure TFrmDialog1.CheckBoxNetZayavkiPropertiesEditValueChanged(
  Sender: TObject);
begin
  if CheckBoxNetZayavki.Checked=true then
    begin
      LabelZayavkaNum.Enabled:=false;
      LabelDateEdit4.Enabled:=false;
      TextEditZayavkaNum.Enabled:=false;
      TextEditZayavkaNum.Text:='0';
      DateEdit4.Enabled:=false;
      DateEdit4.Date:=Date;
    end
  else
    begin
      LabelZayavkaNum.Enabled:=true;
      LabelDateEdit4.Enabled:=true;
      TextEditZayavkaNum.Enabled:=true;
      TextEditZayavkaNum.Text:='';
      DateEdit4.Enabled:=true;
      DateEdit4.Date:=Date;
    end;
end;

procedure TFrmDialog1.CheckBoxShtukPropertiesEditValueChanged(
  Sender: TObject);
begin
  if CheckBoxShtuk.Checked=true then
    CalcEditShtuk.Enabled:=true
  else
    begin
      CalcEditShtuk.Enabled:=false;
      CalcEditShtuk.EditValue:=0;
    end;
end;

procedure TFrmDialog1.ButtonNazvDokClick(Sender: TObject);
begin
  TblRezh:=1;
  TmpObj:=ComboBoxNazvDok;
  FrmMainWindow.ButtonShowDocVidsClick(sender);
end;

procedure TFrmDialog1.ButtonNazvDokExit(Sender: TObject);
begin
  ADODataSet13.Active:=false;
  ADODataSet13.Active:=true;
end;

procedure TFrmDialog1.ComboBoxNazvDokPropertiesEditValueChanged(
  Sender: TObject);
begin
  hint:=ComboBoxNazvDok.EditText;
end;

procedure TFrmDialog1.ComboBoxUprNadzPropertiesEditValueChanged(
  Sender: TObject);
begin
  hint:=ComboBoxUprNadz.EditText;
end;

procedure TFrmDialog1.cxButton5Click(Sender: TObject);
var
  i:integer;
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
      Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('IDNASPUNKT').Index].Caption:='Код населённого пункта';
      Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('OBLNAME').Index].Caption:='Область';
      Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('RAYONNAME').Index].Caption:='Район';
      Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Caption:='Тип населённого пункта';
      Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTNAME').Index].Caption:='Название населённого пункта';
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
          TcxImageComboBoxProperties(Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items[0].Description:='деревня';
          TcxImageComboBoxProperties(Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items[0].Value:='деревня';
          TcxImageComboBoxProperties(Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items[1].Description:='посёлок';
          TcxImageComboBoxProperties(Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items[1].Value:='посёлок';
          TcxImageComboBoxProperties(Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items[2].Description:='город';
          TcxImageComboBoxProperties(Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items[2].Value:='город';
          TcxImageComboBoxProperties(Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items[3].Description:='воинская часть';
          TcxImageComboBoxProperties(Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items[3].Value:='воинская часть';
          TcxImageComboBoxProperties(Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items[3].Description:='хутор';
          TcxImageComboBoxProperties(Columns[FrmProishObrOtech.cxGrid1DBTableView1.GetColumnByFieldName('NASELPUNKTYPE').Index].Properties).Items[3].Value:='хутор';
    end;
  for i:=0 to FrmProishObrOtech.cxGrid1DBTableView1.ColumnCount-1 do FrmProishObrOtech.cxGrid1DBTableView1.Columns[i].Width:=120;
  Screen.Cursor:=crDefault;
  FrmProishObrOtech.ShowModal;
end;

procedure TFrmDialog1.ButtonMemoSizeClick(Sender: TObject);
begin
  if MemoSort.Height=32 then
    begin
      MemoSort.Height:=117;
      MemoSort.Top:=181;
      MemoSort.Width:=583;
      MemoSort.Properties.ScrollBars:=ssBoth;
    end
  else
    begin
      MemoSort.Height:=32;
      MemoSort.Top:=163;
      MemoSort.Width:=565;
      MemoSort.Properties.ScrollBars:=ssNone;
    end;
end;

procedure TFrmDialog1.RadioGroupPartyPolePropertiesEditValueChanged(
  Sender: TObject);
begin
  if RadioGroupPartyPole.ItemIndex=0 then
    begin
      CalcEditEdIzmOtech.Enabled:=true;
      ComboBoxEdIzmOtech.Enabled:=true;
      cxButton7.Enabled:=true;
      LabelKolvoMest.Enabled:=true;
    end
  else
    begin
      CalcEditEdIzmOtech.Enabled:=false;
      ComboBoxEdIzmOtech.Enabled:=false;
      cxButton7.Enabled:=false;
      LabelKolvoMest.Enabled:=false;
      CalcEditEdIzmOtech.EditValue:=0;
      ComboBoxEdIzmOtech.EditValue:=0;
    end;
end;

procedure TFrmDialog1.TextEditSortPropertiesEditValueChanged(
  Sender: TObject);
begin
  if Length(TextEditSort.Text)=0 then
    begin
      CheckBoxSort.Checked:=false;
      TextEditSort.Text:='нет';
    end;
end;

procedure TFrmDialog1.TextEditNazvTranspPropertiesEditValueChanged(
  Sender: TObject);
begin
  if Length(TextEditNazvTransp.Text)=0 then TextEditNazvTransp.Text:='нет данных';
end;

procedure TFrmDialog1.TextEditMestoPropertiesEditValueChanged(
  Sender: TObject);
begin
  if Length(TextEditMesto.Text)=0 then
    begin
      RadioGroupPunkt.ItemIndex:=0;
      TextEditMesto.Text:='нет';
    end;
end;

procedure TFrmDialog1.MemoSortPropertiesEditValueChanged(Sender: TObject);
begin
   if Length(MemoSort.Text)=0 then MemoSort.Text:='нет данных';
end;

procedure TFrmDialog1.TextEditProishOtechPropertiesEditValueChanged(
  Sender: TObject);
begin
   if Length(TextEditProishOtech.Text)=0 then TextEditProishOtech.Text:='нет данных';
end;

procedure TFrmDialog1.ButtonSprProishozdClick(Sender: TObject);
begin
  TblRezh:=1;
  TmpObj:=ComboBoxProishozhd;
  FrmMainWindow.ButtonShowProishozhdClick(Sender);
end;

procedure TFrmDialog1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if FrmToolExcel<>nil then if FrmToolExcel.Showing=true then FrmToolExcel.Close;
  //разблокируем: пишем в базу, что такой-то ПК освободил номер
  FrmMainWindow.ADOCommandLockRegistration.CommandText:='UPDATE SETUP SET SETUP.IDVAL = "0" WHERE (((SETUP.IDSET)=1))';
  FrmMainWindow.ADOCommandLockRegistration.Execute;
end;

procedure TFrmDialog1.ButtonExcelToolClick(Sender: TObject);
begin
   FrmToolExcel.Show;
end;

procedure TFrmDialog1.FormShow(Sender: TObject);
begin
   ButtonExcelTool.Visible:=false;
end;

procedure TFrmDialog1.ButtonSprProishozdExit(Sender: TObject);
begin
  ADODataSet8.Active:=false;
  ADODataSet8.Active:=true;
end;

procedure TFrmDialog1.TextEditNumEtiketkiPropertiesEditValueChanged(
  Sender: TObject);
begin
  if Length(TextEditNumEtiketki.Text)<1 then TextEditNumEtiketki.Text:='0';
end;

procedure TFrmDialog1.DateEdit3PropertiesEditValueChanged(Sender: TObject);
begin
  if Length(DateEdit3.Text)<1 then DateEdit3.Text:=DateToStr(Date());
end;

procedure TFrmDialog1.TextEditDogovorNumPropertiesEditValueChanged(
  Sender: TObject);
begin
  if Length(TextEditDogovorNum.Text)<1 then TextEditDogovorNum.Text:='0';
end;

procedure TFrmDialog1.TextEditZayavkaNumPropertiesEditValueChanged(
  Sender: TObject);
begin
  if Length(TextEditZayavkaNum.Text)<1 then TextEditZayavkaNum.Text:='0';
end;

procedure TFrmDialog1.DateEdit4PropertiesEditValueChanged(Sender: TObject);
begin
  if Length(DateEdit4.Text)<1 then DateEdit4.Text:=DateToStr(Date());
end;

procedure TFrmDialog1.CalcEditVesPropertiesEditValueChanged(
  Sender: TObject);
begin
  if Length(CalcEditVes.Text)<1 then CalcEditVes.Text:='0';
end;

procedure TFrmDialog1.CalcEditEdIzmPropertiesEditValueChanged(
  Sender: TObject);
begin
  if Length(CalcEditEdIzm.Text)<1 then CalcEditEdIzm.Text:='0';
end;

procedure TFrmDialog1.CalcEditKolichPropertiesEditValueChanged(
  Sender: TObject);
begin
  if Length(CalcEditKolich.Text)<1 then CalcEditKolich.Text:='0';
end;

procedure TFrmDialog1.CalcEditShtukPropertiesEditValueChanged(
  Sender: TObject);
begin
  if Length(CalcEditShtuk.Text)<1 then CalcEditShtuk.Text:='0';
end;

procedure TFrmDialog1.cxCalcEdit1PropertiesChange(Sender: TObject);
begin
  if Length(CalcEditVesOtech.Text)<1 then CalcEditVesOtech.Text:='0';
end;

procedure TFrmDialog1.cxCalcEdit2PropertiesEditValueChanged(
  Sender: TObject);
begin
  if Length(CalcEditEdIzmOtech.Text)<1 then CalcEditEdIzmOtech.Text:='0';
end;

procedure TFrmDialog1.RadioGroupNumTypePropertiesChange(Sender: TObject);
begin
   if RadioGroupNumType.ItemIndex=1 then
     begin
       CalcEdit1.Enabled:=true;
       DateEdit1.Date:=Date-1;
     end
   else
     begin
       CalcEdit1.Enabled:=false;
       CalcEdit1.Value:=LastNumProtForAuto;
       DateEdit1.Date:=StrToDate(DateLastDog);
     end;  
   //переправляем тип номерации  
   TypeNumeration:=RadioGroupNumType.ItemIndex;
end;

end.
