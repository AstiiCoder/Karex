unit PeriodOtch;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxCheckListBox, AdvGlowButton, cxTextEdit, cxMaskEdit,
  cxDropDownEdit, cxCalendar, cxLabel, cxControls, cxContainer, cxEdit,
  cxGroupBox, DateUtils, DB, ADODB, cxDBCheckListBox, cxStyles,
  cxCustomData, cxGraphics, cxFilter, cxData, cxDataStorage, cxDBData,
  cxGridLevel, cxClasses, cxGridCustomView, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxGrid, ImgList, cxEditConsts,
  cxCheckBox;

type
  TFrmPeriodOtch = class(TForm)
    cxGroupBox1: TcxGroupBox;
    cxLabel1: TcxLabel;
    Date1: TcxDateEdit;
    LabelDate2: TcxLabel;
    Date2: TcxDateEdit;
    cxGroupBox2: TcxGroupBox;
    cxGroupBox3: TcxGroupBox;
    ButtonOK: TAdvGlowButton;
    ButtonCancel: TAdvGlowButton;
    ADODataSet1: TADODataSet;
    DataSource1: TDataSource;
    ADOCommand1: TADOCommand;
    cxGrid1: TcxGrid;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    ButtonSelectAll: TAdvGlowButton;
    ButtonCancelAll: TAdvGlowButton;
    ButtonEdit: TDBAdvGlowButton;
    ImageList1: TImageList;
    cxLabel2: TcxLabel;
    Date3: TcxDateEdit;
    cxLabel3: TcxLabel;
    Date4: TcxDateEdit;
    CheckBoxNoFoundObj: TcxCheckBox;
    procedure ButtonCancelClick(Sender: TObject);
    procedure ButtonOKClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ButtonSelectAllClick(Sender: TObject);
    procedure ButtonCancelAllClick(Sender: TObject);
    procedure CreateOneSmo;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmPeriodOtch: TFrmPeriodOtch;

implementation

uses
  MainWindow;

{$R *.dfm}

procedure TFrmPeriodOtch.ButtonCancelClick(Sender: TObject);
begin
  PerStart:='';
  close;
end;

procedure TFrmPeriodOtch.ButtonOKClick(Sender: TObject);
var
  LYear, LMonth, LDay: Word;
begin
  DecodeDate(Date1.Date, LYear, LMonth, LDay);
  PerStart:=IntTostr(LMonth)+'/'+IntTostr(LDay)+'/'+IntTostr(LYear);
  DecodeDate(Date2.Date, LYear, LMonth, LDay);
  PerFinish:=IntTostr(LMonth)+'/'+IntTostr(LDay)+'/'+IntTostr(LYear);
  cxGrid1.SetFocus;
  Date2.SetFocus;
  Date1.SetFocus;
  ShowNotFoundObj:=CheckBoxNoFoundObj.Checked;
  close;
end;

procedure TFrmPeriodOtch.FormCreate(Sender: TObject);
begin
  Date1.Date:=StartOfTheMonth(Date()-30);
  Date2.Date:=EndOfTheMonth(Date());
  Date3.Date:=StartOfTheMonth(Date()-30);
  Date4.Date:=EndOfTheMonth(Date());
  cxSetResourceString(@cxSDatePopupToday, 'Сегодня'); //зависимость cxClasses
end;

procedure TFrmPeriodOtch.ButtonSelectAllClick(Sender: TObject);
begin
  ADOCommand1.CommandText:='UPDATE FIELDOTCH SET FIELDOTCH.NEED = True;';
  ADOCommand1.Execute;
  ADODataSet1.Active:=false;
  cxGrid1DBTableView1.ClearItems;
  ADODataSet1.Active:=true;
  CreateOneSmo;
end;

procedure TFrmPeriodOtch.ButtonCancelAllClick(Sender: TObject);
begin
  ADOCommand1.CommandText:='UPDATE FIELDOTCH SET FIELDOTCH.NEED = False;';
  ADOCommand1.Execute;
  ADODataSet1.Active:=false;
  cxGrid1DBTableView1.ClearItems;
  ADODataSet1.Active:=true;
  CreateOneSmo;
end;

procedure TFrmPeriodOtch.CreateOneSmo;
begin
  with cxGrid1DBTableView1.DataController do
    CreateAllItems;
  with cxGrid1DBTableView1 do
    begin
      Columns[cxGrid1DBTableView1.GetColumnByFieldName('NEED').Index].Caption:='Включить';
      Columns[cxGrid1DBTableView1.GetColumnByFieldName('NEED').Index].Width:=70;
      Columns[cxGrid1DBTableView1.GetColumnByFieldName('FIELDNAME').Index].Visible:=false;
      Columns[cxGrid1DBTableView1.GetColumnByFieldName('FIELDNAME').Index].Hidden:=true;
      Columns[cxGrid1DBTableView1.GetColumnByFieldName('FIELDCAPTION').Index].Caption:='Название поля';
      Columns[cxGrid1DBTableView1.GetColumnByFieldName('FIELDCAPTION').Index].Options.Editing:=false;
      Columns[cxGrid1DBTableView1.GetColumnByFieldName('FIELDCAPTION').Index].Width:=320;
    end;
end;

procedure TFrmPeriodOtch.FormShow(Sender: TObject);
begin
  PerStart:='';
  ShowNotFoundObj:=false;
end;

end.
