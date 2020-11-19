unit Dialog2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxLookAndFeelPainters, ExtCtrls, AdvPanel, AdvGlowButton,
  cxCheckBox, cxGroupBox, cxRadioGroup, cxTextEdit, cxDropDownEdit, cxCalc,
  StdCtrls, cxButtons, cxMaskEdit, cxControls, cxContainer, cxEdit,
  cxLabel, ComCtrls, DB, ADODB, cxLookupEdit, cxDBLookupEdit,
  cxDBLookupComboBox, cxGraphics, Menus, cxListView;

type
  TFrmDialog2 = class(TForm)
    DialogWizard: TPageControl;
    TabSheet1: TTabSheet;
    LabelDlgName: TcxLabel;
    TabSheet2: TTabSheet;
    cxLabel1: TcxLabel;
    TabSheet3: TTabSheet;
    LabelNameClient: TcxLabel;
    cxLabel5: TcxLabel;
    TabSheet4: TTabSheet;
    cxLabel6: TcxLabel;
    TabSheet5: TTabSheet;
    cxLabel14: TcxLabel;
    TabSheet6: TTabSheet;
    cxLabel19: TcxLabel;
    TabSheet7: TTabSheet;
    cxLabel20: TcxLabel;
    TabSheet8: TTabSheet;
    cxLabel23: TcxLabel;
    TabSheet9: TTabSheet;
    cxLabel24: TcxLabel;
    TabSheet10: TTabSheet;
    LabelInfo: TcxLabel;
    cxLabel26: TcxLabel;
    ButtonForward: TAdvGlowButton;
    ButtonCancel: TAdvGlowButton;
    ButtonBackward: TAdvGlowButton;
    AdvPanel1: TAdvPanel;
    cxLabel27: TcxLabel;
    cxLabel28: TcxLabel;
    TextEditDolzhnSpets: TcxTextEdit;
    RadioGroupImport: TcxRadioGroup;
    CalcEditProtokolNum: TcxCalcEdit;
    cxRadioGroup3: TcxRadioGroup;
    cxRadioGroup4: TcxRadioGroup;
    cxRadioGroup1: TcxRadioGroup;
    LabelPodskazka: TcxLabel;
    ComboBoxSpetsialist: TcxLookupComboBox;
    ADODataSet1: TADODataSet;
    DataSource1: TDataSource;
    ADODataSet2: TADODataSet;
    DataSource2: TDataSource;
    ADODataSet3: TADODataSet;
    DataSource3: TDataSource;
    ADODataSet4: TADODataSet;
    DataSource4: TDataSource;
    ADOCommand1: TADOCommand;
    ADODataSet5: TADODataSet;
    cxLabel8: TcxLabel;
    TextEditDateProtokol: TcxTextEdit;
    ListViewKarObj: TcxListView;
    ButtonAddKarObj: TAdvGlowButton;
    ButtonDelKarObj: TAdvGlowButton;
    ListViewNekarObjOtsRF: TcxListView;
    AdvGlowButton1: TAdvGlowButton;
    AdvGlowButton2: TAdvGlowButton;
    ListViewNekarObjProchii: TcxListView;
    AdvGlowButton3: TAdvGlowButton;
    AdvGlowButton4: TAdvGlowButton;
    TextEditIDExp: TcxTextEdit;
    ListViewKarObjOtech: TcxListView;
    ListViewNekarObjOtsRFOtech: TcxListView;
    ListViewNekarObjProchiiOtech: TcxListView;
    procedure ButtonForwardClick(Sender: TObject);
    procedure ButtonBackwardClick(Sender: TObject);
    procedure ButtonCancelClick(Sender: TObject);
    procedure cxButton10Click(Sender: TObject);
    procedure ButtonSprTransportClick(Sender: TObject);
    procedure ButtonSprTransportExit(Sender: TObject);
    procedure ButtonSprNekarObjClick(Sender: TObject);
    procedure ButtonSprNekarObjExit(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure CalcEditProtokolNumPropertiesEditValueChanged(
      Sender: TObject);
    procedure ComboBoxSpetsialistPropertiesEditValueChanged(
      Sender: TObject);
    procedure ButtonAddKarObjClick(Sender: TObject);
    procedure ButtonDelKarObjClick(Sender: TObject);
    procedure AdvGlowButton2Click(Sender: TObject);
    procedure AdvGlowButton4Click(Sender: TObject);
    procedure AdvGlowButton1Click(Sender: TObject);
    procedure AdvGlowButton3Click(Sender: TObject);
    procedure RadioGroupImportPropertiesEditValueChanged(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmDialog2: TFrmDialog2;
  ObjName, ObjSost, ObjPrim: string;
  NumExp: integer;
  TmpObj: TcxLookupComboBox;

implementation

uses
  MainWindow, AddObjToExp, Grid1, AddObjToExpOtech;

{$R *.dfm}

procedure TFrmDialog2.ButtonForwardClick(Sender: TObject);
var
  SQLstr:string;
  i: integer;
begin
  if DialogWizard.ActivePage = TabSheet1 then
    begin
      DialogWizard.ActivePage := TabSheet2;
      ButtonBackward.Enabled:=true;
      TabSheet2.Show;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet2 then
    begin
      DialogWizard.ActivePage := TabSheet3;
      ButtonForward.Enabled:=false;
      CalcEditProtokolNum.EditText:='';
      TabSheet3.Show;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet3 then
    begin
      //запись в реестр экспертиз
      SQLstr:='INSERT INTO EXPERTIZE_REESTR (ISIMPORT, NUMPROTOKOL, SPETSIALIST)';
      SQLstr:=SQLstr+' SELECT '+IntToStr(RadioGroupImport.ItemIndex)+' AS B1, '+TextEditIDExp.Text+' AS B2, '+IntToStr(ComboBoxSpetsialist.EditValue)+' AS B3';
      ADOCommand1.CommandText:=SQLstr;
      ADOCommand1.Execute;
      //определение текущего номера экспертизы
      ADODataSet5.Active:=false;
      SQLstr:='SELECT Max(IDEXPERTIZE) AS M FROM EXPERTIZE_REESTR WHERE (((SPETSIALIST)='+IntToStr(ComboBoxSpetsialist.EditValue)+') AND ((ISIMPORT)='+IntToStr(RadioGroupImport.ItemIndex)+'))';
      ADODataSet5.CommandText:=SQLstr;
      ADODataSet5.Active:=true;
      ADODataSet5.First;
      NumExp:=ADODataSet5.FieldByName('M').AsInteger;
      Caption:='Текущий номер экспертизы: '+IntToStr(NumExp);
      DialogWizard.ActivePage := TabSheet4;
      TabSheet4.Show;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet4 then
    begin
      if cxRadioGroup3.ItemIndex=1 then DialogWizard.ActivePage := TabSheet5 else DialogWizard.ActivePage := TabSheet6;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet5 then
    begin
      if ((ListViewKarObj.Items.Count=0) and (RadioGroupImport.ItemIndex=0)) or ((ListViewKarObjOtech.Items.Count=0) and (RadioGroupImport.ItemIndex=1)) then
        begin
          showmessage('Укажите обнаруженные объекты или выберите "объекты не обнаружены!"');
          exit;
        end;
      DialogWizard.ActivePage := TabSheet6;
      TabSheet6.Show;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet6 then
    begin
      if cxRadioGroup4.ItemIndex=1 then DialogWizard.ActivePage := TabSheet7 else DialogWizard.ActivePage := TabSheet8;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet7 then
    begin
      if ((ListViewNekarObjOtsRF.Items.Count=0) and (RadioGroupImport.ItemIndex=0)) or ((ListViewNekarObjOtsRFOtech.Items.Count=0) and (RadioGroupImport.ItemIndex=1)) then
        begin
          showmessage('Укажите обнаруженные объекты или выберите "объекты не обнаружены!"');
          exit;
        end;
      DialogWizard.ActivePage := TabSheet8;
      TabSheet8.Show;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet8 then
    begin
      if cxRadioGroup1.ItemIndex=1 then
        DialogWizard.ActivePage := TabSheet9
      else
        begin
          DialogWizard.ActivePage := TabSheet10;
          ButtonBackward.Enabled:=false;
          ButtonForward.Caption:='Готово!';
          ButtonCancel.Enabled:=false;
        end;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet9 then
    begin
      if ((ListViewNekarObjProchii.Items.Count=0) and (RadioGroupImport.ItemIndex=0)) or ((ListViewNekarObjProchiiOtech.Items.Count=0) and (RadioGroupImport.ItemIndex=1)) then
        begin
          showmessage('Укажите обнаруженные объекты или выберите "объекты не обнаружены!"');
          exit;
        end;
      DialogWizard.ActivePage := TabSheet10;
      TabSheet10.Show;
      ButtonBackward.Enabled:=false;
      ButtonForward.Caption:='Готово!';
      ButtonCancel.Enabled:=false;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet10 then
    begin
        //запись в базу
        if RadioGroupImport.ItemIndex=0 then
          begin
            if ListViewKarObj.Items.Count>0 then
              for i:=1 to ListViewKarObj.Items.Count do
                begin
            SQLstr:='INSERT INTO EXPERTIZE_DATA (NUMEXPERTIZE, TYPEEXPERTIZE, TYPEOBJ, CODEOBJ, SOSTOYANIE, PRIMECH)';
            SQLstr:=SQLstr+' SELECT '+IntToStr(NumExp)+' AS NUM, "'+ExpTypeName+'" AS TYPEEXP, 0 AS TYPEOBJ, '+ListViewKarObj.Items[i-1].SubItems[2]+' AS CODEOBJ, "'+ListViewKarObj.Items[i-1].SubItems[0]+'" AS SOST, "'+ListViewKarObj.Items[i-1].SubItems[1]+'" AS PRIMECH;';
            ADOCommand1.CommandText:=SQLstr;
            ADOCommand1.Execute;
                end;
            if ListViewNekarObjOtsRF.Items.Count>0 then
              for i:=1 to ListViewNekarObjOtsRF.Items.Count do
                begin
            SQLstr:='INSERT INTO EXPERTIZE_DATA (NUMEXPERTIZE, TYPEEXPERTIZE, TYPEOBJ, CODEOBJ, SOSTOYANIE, PRIMECH)';
            SQLstr:=SQLstr+' SELECT '+IntToStr(NumExp)+' AS NUM, "'+ExpTypeName+'" AS TYPEEXP, 1 AS TYPEOBJ, '+ListViewNekarObjOtsRF.Items[i-1].SubItems[2]+' AS CODEOBJ, "'+ListViewNekarObjOtsRF.Items[i-1].SubItems[0]+'" AS SOST, "'+ListViewNekarObjOtsRF.Items[i-1].SubItems[1]+'" AS PRIMECH;';
            ADOCommand1.CommandText:=SQLstr;
            ADOCommand1.Execute;
                end;
            if ListViewNekarObjProchii.Items.Count>0 then
              for i:=1 to ListViewNekarObjProchii.Items.Count do
                begin
            SQLstr:='INSERT INTO EXPERTIZE_DATA (NUMEXPERTIZE, TYPEEXPERTIZE, TYPEOBJ, CODEOBJ, SOSTOYANIE, PRIMECH)';
            SQLstr:=SQLstr+' SELECT '+IntToStr(NumExp)+' AS NUM, "'+ExpTypeName+'" AS TYPEEXP, 2 AS TYPEOBJ, '+ListViewNekarObjProchii.Items[i-1].SubItems[2]+' AS CODEOBJ, "'+ListViewNekarObjProchii.Items[i-1].SubItems[0]+'" AS SOST, "'+ListViewNekarObjProchii.Items[i-1].SubItems[1]+'" AS PRIMECH;';
            ADOCommand1.CommandText:=SQLstr;
            ADOCommand1.Execute;
                end;
          end
        else
          begin
            if ListViewKarObjOtech.Items.Count>0 then
              for i:=1 to ListViewKarObjOtech.Items.Count do
                begin
            SQLstr:='INSERT INTO EXPERTIZE_DATA (NUMEXPERTIZE, TYPEEXPERTIZE, TYPEOBJ, CODEOBJ, SOSTOYANIE, PRIMECH)';
            SQLstr:=SQLstr+' SELECT '+IntToStr(NumExp)+' AS NUM, "'+ExpTypeName+'" AS TYPEEXP, 0 AS TYPEOBJ, '+ListViewKarObjOtech.Items[i-1].SubItems[2]+' AS CODEOBJ, "'+ListViewKarObjOtech.Items[i-1].SubItems[0]+'" AS SOST, "'+ListViewKarObjOtech.Items[i-1].SubItems[1]+'" AS PRIMECH;';
            ADOCommand1.CommandText:=SQLstr;
            ADOCommand1.Execute;
                end;
            if ListViewNekarObjOtsRFOtech.Items.Count>0 then
              for i:=1 to ListViewNekarObjOtsRFOtech.Items.Count do
                begin
            SQLstr:='INSERT INTO EXPERTIZE_DATA (NUMEXPERTIZE, TYPEEXPERTIZE, TYPEOBJ, CODEOBJ, SOSTOYANIE, PRIMECH)';
            SQLstr:=SQLstr+' SELECT '+IntToStr(NumExp)+' AS NUM, "'+ExpTypeName+'" AS TYPEEXP, 1 AS TYPEOBJ, '+ListViewNekarObjOtsRFOtech.Items[i-1].SubItems[2]+' AS CODEOBJ, "'+ListViewNekarObjOtsRFOtech.Items[i-1].SubItems[0]+'" AS SOST, "'+ListViewNekarObjOtsRFOtech.Items[i-1].SubItems[1]+'" AS PRIMECH;';
            ADOCommand1.CommandText:=SQLstr;
            ADOCommand1.Execute;
                end;
            if ListViewNekarObjProchiiOtech.Items.Count>0 then
              for i:=1 to ListViewNekarObjProchiiOtech.Items.Count do
                begin
            SQLstr:='INSERT INTO EXPERTIZE_DATA (NUMEXPERTIZE, TYPEEXPERTIZE, TYPEOBJ, CODEOBJ, SOSTOYANIE, PRIMECH)';
            SQLstr:=SQLstr+' SELECT '+IntToStr(NumExp)+' AS NUM, "'+ExpTypeName+'" AS TYPEEXP, 2 AS TYPEOBJ, '+ListViewNekarObjProchiiOtech.Items[i-1].SubItems[2]+' AS CODEOBJ, "'+ListViewNekarObjProchiiOtech.Items[i-1].SubItems[0]+'" AS SOST, "'+ListViewNekarObjProchiiOtech.Items[i-1].SubItems[1]+'" AS PRIMECH;';
            ADOCommand1.CommandText:=SQLstr;
            ADOCommand1.Execute;
                end;
          end;
        ListViewKarObj.Clear;
        ListViewNekarObjOtsRF.Clear;
        ListViewNekarObjProchii.Clear;
        ButtonForward.Caption:='Продолжить >';
        ButtonForward.Enabled:=true;
        ButtonBackward.Enabled:=false;
        ButtonCancel.Enabled:=true;
        DialogWizard.ActivePage := TabSheet1;
        Caption:='Мастер проведения карантинной экспертизы, следующая экспертиза ';
    end;
end;

procedure TFrmDialog2.ButtonBackwardClick(Sender: TObject);
begin
  if DialogWizard.ActivePage = TabSheet9 then
    begin
      DialogWizard.ActivePage := TabSheet8;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet8 then
    begin
      if cxRadioGroup4.ItemIndex=1 then DialogWizard.ActivePage := TabSheet7 else DialogWizard.ActivePage := TabSheet6;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet7 then
    begin
      DialogWizard.ActivePage := TabSheet6;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet6 then
    begin
      if cxRadioGroup3.ItemIndex=1 then DialogWizard.ActivePage := TabSheet5 else DialogWizard.ActivePage := TabSheet4;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet5 then
    begin
      DialogWizard.ActivePage := TabSheet4;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet4 then
    begin
      DialogWizard.ActivePage := TabSheet3;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet3 then
    begin
      DialogWizard.ActivePage := TabSheet2;
      ButtonForward.Enabled:=true;
      exit;
    end;
  if DialogWizard.ActivePage = TabSheet2 then
    begin
      DialogWizard.ActivePage := TabSheet1;
      ButtonBackward.Enabled:=false;
      exit;
    end;
end;

procedure TFrmDialog2.ButtonCancelClick(Sender: TObject);
begin
  close;
end;

procedure TFrmDialog2.cxButton10Click(Sender: TObject);
begin
  FrmMainWindow.ButtonShowClientsClick(Sender);
end;

procedure TFrmDialog2.ButtonSprTransportClick(Sender: TObject);
begin
  TblRezh:=1;
  FrmMainWindow.ButtonShowTransportClick(Sender);
end;

procedure TFrmDialog2.ButtonSprTransportExit(Sender: TObject);
begin
  ADODataSet3.Refresh;
end;

procedure TFrmDialog2.ButtonSprNekarObjClick(Sender: TObject);
begin
  TblRezh:=1;
  FrmMainWindow.ButtonNekarObjProcheeClick(Sender);
end;

procedure TFrmDialog2.ButtonSprNekarObjExit(Sender: TObject);
begin
  ADODataSet3.Refresh;
end;

procedure TFrmDialog2.FormCreate(Sender: TObject);
begin
  ADODataSet1.Active:=true;
  ADODataSet2.Active:=true;
  ADODataSet3.Active:=true;
  ADODataSet4.Active:=true;
  ListViewNekarObjProchii.Top:=54;
  ListViewNekarObjProchiiOtech.Top:=54;
  ListViewNekarObjOtsRF.Top:=54;
  ListViewNekarObjOtsRFOtech.Top:=54;
  ListViewKarObj.Top:=54;
  ListViewKarObjOtech.Top:=54;
  ListViewNekarObjProchii.Height:=182;
  ListViewNekarObjProchiiOtech.Height:=182;
  ListViewNekarObjOtsRF.Height:=182;
  ListViewNekarObjOtsRFOtech.Height:=182;
  ListViewKarObj.Height:=182;
  ListViewKarObjOtech.Height:=182;
end;

procedure TFrmDialog2.CalcEditProtokolNumPropertiesEditValueChanged(
  Sender: TObject);
var
  SQLstr:string;
begin
  ADODataSet5.Active:=false;
  //ADODataSet5.CommandText:='SELECT NUMPROTOKOL, DATEPROTOKOL FROM REGISTRATION WHERE (((REGISTRATION.NUMPROTOKOL)='+IntToStr(CalcEditProtokolNum.EditValue)+')) ORDER BY DATEPROTOKOL DESC;';
  SQLstr:='SELECT NUMPROTOKOL, DATEPROTOKOL, EXPCLOSE, IDREGISTRATION FROM REGISTRATION WHERE (((REGISTRATION.NUMPROTOKOL)='+IntToStr(CalcEditProtokolNum.EditValue);
  SQLstr:=SQLstr+') AND ((REGISTRATION.ISIMPORT)='+IntToStr(RadioGroupImport.ItemIndex)+')) ORDER BY DATEPROTOKOL DESC;';
  ADODataSet5.CommandText:=SQLstr;
  ADODataSet5.Active:=true;
  if ADODataSet5.RecordCount>0 then
    begin
      ADODataSet5.First;
      TextEditDateProtokol.Text:=ADODataSet5.FieldByName('DATEPROTOKOL').AsString;
      TextEditIDExp.Text:=ADODataSet5.FieldByName('IDREGISTRATION').AsString;
      if StrToBool(ADODataSet5.FieldByName('EXPCLOSE').AsString)=false then
        ButtonForward.Enabled:=true
      else
        begin
           ButtonForward.Enabled:=false;
           showmessage('Ввод данных экпертизы по данному протоколу закончен!'+Chr(13)+'Обратитесь в организационный отдел.');
           exit;
        end;
    end
  else
    begin
      TextEditDateProtokol.EditingText:='';
      ButtonForward.Enabled:=false;
      if CalcEditProtokolNum.EditValue<>0 then showmessage('Протокол с выбранным номером не зарегистрирован!');
      exit;
    end;
  if RadioGroupImport.ItemIndex=1 then
    begin
      ADODataSet5.Active:=false;
      SQLstr:='SELECT SORTTMPROCH FROM OTECHMEMOSORT WHERE (((OTECHMEMOSORT.IDREGISTRATION)='+TextEditIDExp.Text+'))';
      ADODataSet5.CommandText:=SQLstr;
      ADODataSet5.Active:=true;
      if ADODataSet5.RecordCount>0 then FrmAddObjToExpOtech.MemoOchag.Text:=ADODataSet5.FieldByName('SORTTMPROCH').AsString;
    end;
end;

procedure TFrmDialog2.ComboBoxSpetsialistPropertiesEditValueChanged(
  Sender: TObject);
begin
  //определение должности по фамилии
  ADODataSet5.Active:=false;
  ADODataSet5.CommandText:='SELECT DOLZHNSPETSIALIST FROM SPETSIALISTS_LAB WHERE (((IDSPETSIALIST)='+IntToStr(ComboBoxSpetsialist.EditValue)+'));';
  ADODataSet5.Active:=true;
  if ADODataSet5.RecordCount>0 then
    begin
      ADODataSet5.First;
      TextEditDolzhnSpets.Text:=ADODataSet5.FieldByName('DOLZHNSPETSIALIST').AsString;
    end
  else
    TextEditDolzhnSpets.Text:='';
end;

procedure TFrmDialog2.ButtonAddKarObjClick(Sender: TObject);
var
  ListItem: TListItem;
begin
  if RadioGroupImport.ItemIndex=0 then
    begin
      ObjName:='';
      FrmAddObjToExp.ADODataSet2.Active:=false;
      FrmAddObjToExp.ADODataSet2.CommandText:='select * from KAROBJECTS WHERE (((KAROBJECTS.TYPEOBJ)="'+ExpTypeName+'")) ORDER BY TYPEOBJ, NAMEOBJ;';
      FrmAddObjToExp.ADODataSet2.Active:=true;
      if FrmAddObjToExp= NIL then FrmAddObjToExp := TFrmAddObjToExp.Create(owner);
      FrmAddObjToExp.ShowModal;
      if ObjName<>'' then
        begin
            begin
              ListItem := ListViewKarObj.Items.Add;
              ListItem.Caption := ObjName;
              ListItem.SubItems.Add(ObjSost);
              ListItem.SubItems.Add(ObjPrim);
              ListItem.SubItems.Add(FrmAddObjToExp.ComboBoxObjName.EditValue);
            end;
        end;
    end
  else
    begin
      ObjName:='';
      FrmAddObjToExpOtech.ADODataSet2.Active:=false;
      FrmAddObjToExpOtech.ADODataSet2.CommandText:='select * from KAROBJECTS WHERE (((KAROBJECTS.TYPEOBJ)="'+ExpTypeName+'")) ORDER BY TYPEOBJ, NAMEOBJ;';
      FrmAddObjToExpOtech.ADODataSet2.Active:=true;
      if FrmAddObjToExpOtech= NIL then FrmAddObjToExpOtech := TFrmAddObjToExpOtech.Create(owner);
      FrmAddObjToExpOtech.ShowModal;
      if ObjName<>'' then
        begin
            begin
              ListItem := ListViewKarObjOtech.Items.Add;
              ListItem.Caption := ObjName;
              ListItem.SubItems.Add(ObjSost);
              ListItem.SubItems.Add(ObjPrim);
              ListItem.SubItems.Add(FrmAddObjToExpOtech.ComboBoxObjName.EditValue);
            end;
        end;
    end;
end;

procedure TFrmDialog2.ButtonDelKarObjClick(Sender: TObject);
begin
  if RadioGroupImport.ItemIndex=0 then
    try
      ListViewKarObj.Selected.Delete;
    except
      showmessage('Выделите строку для удаления!');
    end
  else
    try
      ListViewKarObjOtech.Selected.Delete;
    except
      showmessage('Выделите строку для удаления!');
    end
end;

procedure TFrmDialog2.AdvGlowButton2Click(Sender: TObject);
begin
  if RadioGroupImport.ItemIndex=0 then
    try
      ListViewNekarObjOtsRF.Selected.Delete;
    except
      showmessage('Выделите строку для удаления!');
    end
  else
    try
      ListViewNekarObjOtsRFOtech.Selected.Delete;
    except
      showmessage('Выделите строку для удаления!');
    end
end;

procedure TFrmDialog2.AdvGlowButton4Click(Sender: TObject);
begin
  if RadioGroupImport.ItemIndex=0 then
    try
      ListViewNekarObjProchii.Selected.Delete;
    except
      showmessage('Выделите строку для удаления!');
    end
  else
    try
      ListViewNekarObjProchiiOtech.Selected.Delete;
    except
      showmessage('Выделите строку для удаления!');
    end;
end;

procedure TFrmDialog2.AdvGlowButton1Click(Sender: TObject);
var
  ListItem: TListItem;
begin
  if RadioGroupImport.ItemIndex=0 then
    begin
      ObjName:='';
      FrmAddObjToExp.ADODataSet2.Active:=false;
      FrmAddObjToExp.ADODataSet2.CommandText:='select IDNEKAROBJ AS IDOBJ, NAMELAT AS NAMEOBJ, NAMERUS AS NAMEOBJRUS from NEKAROBJECTS_NEZARRF WHERE (((NEKAROBJECTS_NEZARRF.TYPENEKAROBJ)="'+ExpTypeName+'")) ORDER BY TYPENEKAROBJ, NAMELAT;';
      FrmAddObjToExp.ADODataSet2.Active:=true;
      if FrmAddObjToExp= NIL then FrmAddObjToExp := TFrmAddObjToExp.Create(owner);
      FrmAddObjToExp.ShowModal;
      if ObjName<>'' then
          begin
            ListItem := ListViewNekarObjOtsRF.Items.Add;
            ListItem.Caption := ObjName;
            ListItem.SubItems.Add(ObjSost);
            ListItem.SubItems.Add(ObjPrim);
            ListItem.SubItems.Add(FrmAddObjToExp.ComboBoxObjName.EditValue);
          end;
    end
  else
    begin
      ObjName:='';
      FrmAddObjToExpOtech.ADODataSet2.Active:=false;
      FrmAddObjToExpOtech.ADODataSet2.CommandText:='select IDNEKAROBJ AS IDOBJ, NAMELAT AS NAMEOBJ, NAMERUS AS NAMEOBJRUS from NEKAROBJECTS_NEZARRF WHERE (((NEKAROBJECTS_NEZARRF.TYPENEKAROBJ)="'+ExpTypeName+'")) ORDER BY TYPENEKAROBJ, NAMELAT;';
      FrmAddObjToExpOtech.ADODataSet2.Active:=true;
      if FrmAddObjToExpOtech= NIL then FrmAddObjToExpOtech := TFrmAddObjToExpOtech.Create(owner);
      FrmAddObjToExpOtech.ShowModal;
      if ObjName<>'' then
        begin
              ListItem := ListViewNekarObjOtsRFOtech.Items.Add;
              ListItem.Caption := ObjName;
              ListItem.SubItems.Add(ObjSost);
              ListItem.SubItems.Add(ObjPrim);
              ListItem.SubItems.Add(FrmAddObjToExpOtech.ComboBoxObjName.EditValue);
        end;
    end;
end;

procedure TFrmDialog2.AdvGlowButton3Click(Sender: TObject);
var
  ListItem: TListItem;
begin
  if RadioGroupImport.ItemIndex=0 then
    begin
      ObjName:='';
      FrmAddObjToExp.ADODataSet2.Active:=false;
      FrmAddObjToExp.ADODataSet2.CommandText:='select IDNEKAROBJ AS IDOBJ, NAMELAT AS NAMEOBJ, NAMERUS AS NAMEOBJRUS from NEKAROBJECTS_PROCHIE WHERE (((NEKAROBJECTS_PROCHIE.TYPENEKAROBJ)="'+ExpTypeName+'")) ORDER BY TYPENEKAROBJ, NAMELAT;';
      FrmAddObjToExp.ADODataSet2.Active:=true;
      if FrmAddObjToExp= NIL then FrmAddObjToExp := TFrmAddObjToExp.Create(owner);
      FrmAddObjToExp.ShowModal;
      if ObjName<>'' then
        begin
          ListItem := ListViewNekarObjProchii.Items.Add;
          ListItem.Caption := ObjName;
          ListItem.SubItems.Add(ObjSost);
          ListItem.SubItems.Add(ObjPrim);
          ListItem.SubItems.Add(FrmAddObjToExp.ComboBoxObjName.EditValue);
        end;
    end
  else
    begin
      ObjName:='';
      FrmAddObjToExpOtech.ADODataSet2.Active:=false;
      FrmAddObjToExpOtech.ADODataSet2.CommandText:='select IDNEKAROBJ AS IDOBJ, NAMELAT AS NAMEOBJ, NAMERUS AS NAMEOBJRUS from NEKAROBJECTS_PROCHIE WHERE (((NEKAROBJECTS_PROCHIE.TYPENEKAROBJ)="'+ExpTypeName+'")) ORDER BY TYPENEKAROBJ, NAMELAT;';
      FrmAddObjToExpOtech.ADODataSet2.Active:=true;
      if FrmAddObjToExpOtech= NIL then FrmAddObjToExpOtech := TFrmAddObjToExpOtech.Create(owner);
      FrmAddObjToExpOtech.ShowModal;
      if ObjName<>'' then
        begin
              ListItem := ListViewNekarObjProchiiOtech.Items.Add;
              ListItem.Caption := ObjName;
              ListItem.SubItems.Add(ObjSost);
              ListItem.SubItems.Add(ObjPrim);
              ListItem.SubItems.Add(FrmAddObjToExpOtech.ComboBoxObjName.EditValue);
        end;
    end;
end;

procedure TFrmDialog2.RadioGroupImportPropertiesEditValueChanged(
  Sender: TObject);
begin
  if RadioGroupImport.ItemIndex=0 then
    begin
      ListViewNekarObjProchii.Visible:=true;
      ListViewNekarObjProchiiOtech.Visible:=false;
      ListViewNekarObjOtsRF.Visible:=true;
      ListViewNekarObjOtsRFOtech.Visible:=false;
      ListViewKarObj.Visible:=true;
      ListViewKarObjOtech.Visible:=false;
    end
  else
    begin
      ListViewNekarObjProchii.Visible:=false;
      ListViewNekarObjProchiiOtech.Visible:=true;
      ListViewNekarObjOtsRF.Visible:=false;
      ListViewNekarObjOtsRFOtech.Visible:=true;
      ListViewKarObj.Visible:=false;
      ListViewKarObjOtech.Visible:=true;
    end;
end;

end.
