unit Period;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, AdvGlowButton, cxLabel, cxControls, cxContainer, cxEdit,  cxClasses, cxEditConsts,
  cxTextEdit, cxMaskEdit, cxDropDownEdit, cxCalendar, DateUtils, cxGroupBox,
  cxLookAndFeelPainters;

type
  TFrmPeriod = class(TForm)
    ButtonOK: TAdvGlowButton;
    AdvGlowButton1: TAdvGlowButton;
    cxGroupBox1: TcxGroupBox;
    Date1: TcxDateEdit;
    LabelDate2: TcxLabel;
    Date2: TcxDateEdit;
    cxLabel1: TcxLabel;
    procedure AdvGlowButton1Click(Sender: TObject);
    procedure ButtonOKClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmPeriod: TFrmPeriod;

implementation

uses
  MainWindow;

{$R *.dfm}

procedure TFrmPeriod.AdvGlowButton1Click(Sender: TObject);
begin
  close;
end;

procedure TFrmPeriod.ButtonOKClick(Sender: TObject);
var
  LYear, LMonth, LDay: Word;
begin
  Date1.SetFocus;
  Date2.SetFocus;
  DecodeDate(Date1.Date, LYear, LMonth, LDay);
  PerStart:=IntTostr(LMonth)+'/'+IntTostr(LDay)+'/'+IntTostr(LYear);
  DecodeDate(Date2.Date, LYear, LMonth, LDay);
  PerFinish:=IntTostr(LMonth)+'/'+IntTostr(LDay)+'/'+IntTostr(LYear);
  close;
end;

procedure TFrmPeriod.FormCreate(Sender: TObject);
begin
  Date1.Date:=StartOfTheMonth(Date()-30);
  Date2.Date:=EndOfTheMonth(Date());
  cxSetResourceString(@cxSDatePopupToday, 'Сегодня'); //зависимость cxClasses
end;

procedure TFrmPeriod.FormShow(Sender: TObject);
begin
  PerStart:='';
end;

end.
