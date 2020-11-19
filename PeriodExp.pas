unit PeriodExp;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, AdvGlowButton, cxLabel, cxControls, cxContainer, cxEdit,
  cxTextEdit, cxMaskEdit, cxDropDownEdit, cxCalendar, cxGroupBox,
  cxRadioGroup, DateUtils,  cxClasses, cxEditConsts, cxCheckBox, cxCalc;

type
  TFrmPeriodExp = class(TForm)
    ButtonOK: TAdvGlowButton;
    ButtonCancel: TAdvGlowButton;
    cxGroupBox1: TcxGroupBox;
    cxLabel1: TcxLabel;
    Date1: TcxDateEdit;
    LabelDate2: TcxLabel;
    Date2: TcxDateEdit;
    RadioGroupParamVibor: TcxRadioGroup;
    RadioGroupObj: TcxRadioGroup;
    RadioGroupImport: TcxRadioGroup;
    CalcEditCode: TcxCalcEdit;
    procedure ButtonCancelClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ButtonOKClick(Sender: TObject);
    procedure RadioGroupParamViborPropertiesEditValueChanged(
      Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure RadioGroupParamViborPropertiesChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmPeriodExp: TFrmPeriodExp;

implementation

uses
  MainWindow;

{$R *.dfm}

procedure TFrmPeriodExp.ButtonCancelClick(Sender: TObject);
begin
  PerStart:='';
  close;
end;

procedure TFrmPeriodExp.FormCreate(Sender: TObject);
begin
  Date1.Date:=StartOfTheMonth(Date()-30);
  Date2.Date:=EndOfTheMonth(Date());
  cxSetResourceString(@cxSDatePopupToday, 'Сегодня'); //зависимость cxClasses
end;

procedure TFrmPeriodExp.ButtonOKClick(Sender: TObject);
var
  LYear, LMonth, LDay: Word;
begin
  DecodeDate(Date1.Date, LYear, LMonth, LDay);
  PerStart:=IntTostr(LMonth)+'/'+IntTostr(LDay)+'/'+IntTostr(LYear);
  DecodeDate(Date2.Date, LYear, LMonth, LDay);
  PerFinish:=IntTostr(LMonth)+'/'+IntTostr(LDay)+'/'+IntTostr(LYear);
  Date1.SetFocus;
  close;
end;

procedure TFrmPeriodExp.RadioGroupParamViborPropertiesEditValueChanged(
  Sender: TObject);
begin
  if RadioGroupParamVibor.ItemIndex=1 then width:=563 else width:=311;
  if RadioGroupParamVibor.ItemIndex=2 then CalcEditCode.Visible:=true else CalcEditCode.Visible:=false;
end;

procedure TFrmPeriodExp.FormShow(Sender: TObject);
begin
  RadioGroupParamVibor.ItemIndex:=0;
  if RadioGroupParamVibor.ItemIndex=2 then CalcEditCode.Visible:=true;
  CalcEditCode.Visible:=false;
end;

procedure TFrmPeriodExp.RadioGroupParamViborPropertiesChange(
  Sender: TObject);
begin
  if RadioGroupParamVibor.ItemIndex=1 then width:=563 else width:=311;
  if RadioGroupParamVibor.ItemIndex=2 then CalcEditCode.Visible:=true else CalcEditCode.Visible:=false;
end;

end.
