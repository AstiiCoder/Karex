unit AddObjToExpOtech;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxGraphics, Menus, cxLookAndFeelPainters, cxTextEdit, StdCtrls,
  cxButtons, cxMaskEdit, cxDropDownEdit, cxLookupEdit, cxDBLookupEdit,
  cxDBLookupComboBox, cxControls, cxContainer, cxEdit, cxLabel,
  AdvGlowButton, DB, ADODB, cxMemo, cxCalc;

type
  TFrmAddObjToExpOtech = class(TForm)
    ButtonOK: TAdvGlowButton;
    AdvGlowButton1: TAdvGlowButton;
    cxLabel2: TcxLabel;
    ComboBoxObjName: TcxLookupComboBox;
    cxLabel3: TcxLabel;
    cxLabel1: TcxLabel;
    SostMemo: TcxMemo;
    MemoPrim: TcxMemo;
    ButtonClearAll: TAdvGlowButton;
    ADODataSet2: TADODataSet;
    DataSource2: TDataSource;
    LabelOchag: TcxLabel;
    MemoOchag: TcxMemo;
    cxLabel13: TcxLabel;
    CalcEditEdIzm: TcxCalcEdit;
    cxLabel18: TcxLabel;
    ComboBoxEdIzm: TcxLookupComboBox;
    ADODataSet10: TADODataSet;
    DataSource10: TDataSource;
    ButtonMemoSize: TcxButton;
    procedure ButtonOKClick(Sender: TObject);
    procedure AdvGlowButton1Click(Sender: TObject);
    procedure ButtonClearAllClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ButtonMemoSizeClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmAddObjToExpOtech: TFrmAddObjToExpOtech;
  TmpObj: TcxLookupComboBox;

implementation

uses
  MainWindow, Dialog2;

{$R *.dfm}

procedure TFrmAddObjToExpOtech.ButtonOKClick(Sender: TObject);
begin
  ObjName:=ComboBoxObjName.EditText;
  ObjSost:=SostMemo.Text;
  ObjPrim:=MemoPrim.Text;
  close;
end;

procedure TFrmAddObjToExpOtech.AdvGlowButton1Click(Sender: TObject);
begin
  close;
end;

procedure TFrmAddObjToExpOtech.ButtonClearAllClick(Sender: TObject);
begin
  SostMemo.Text:='';
  MemoPrim.Text:='';
  MemoOchag.Text:='';
  ComboBoxObjName.ItemIndex:=-1;
  CalcEditEdIzm.Value:=0;
  ComboBoxEdIzm.ItemIndex:=-1;
end;

procedure TFrmAddObjToExpOtech.FormShow(Sender: TObject);
begin
  ADODataSet10.Active:=false;
  ADODataSet10.Active:=true;
end;

procedure TFrmAddObjToExpOtech.ButtonMemoSizeClick(Sender: TObject);
begin
  if MemoOchag.Top=125 then
    begin
       MemoOchag.Left:=18;
       MemoOchag.Top:=142;
       MemoOchag.Height:=132;
       MemoOchag.Width:=550;
       LabelOchag.Width:=409;
       MemoOchag.Properties.ScrollBars:=ssBoth;
    end
  else
    begin
       MemoOchag.Left:=144;
       MemoOchag.Top:=125;
       MemoOchag.Height:=47;
       MemoOchag.Width:=406;
       LabelOchag.Width:=129;
       MemoOchag.Properties.ScrollBars:=ssNone;
    end;
end;

end.
