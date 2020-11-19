unit AddObjToExp;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxGraphics, Menus, cxLookAndFeelPainters, cxTextEdit, StdCtrls,
  cxButtons, cxMaskEdit, cxDropDownEdit, cxLookupEdit, cxDBLookupEdit,
  cxDBLookupComboBox, cxControls, cxContainer, cxEdit, cxLabel,
  AdvGlowButton, DB, ADODB, cxMemo;

type
  TFrmAddObjToExp = class(TForm)
    ButtonOK: TAdvGlowButton;
    AdvGlowButton1: TAdvGlowButton;
    cxLabel2: TcxLabel;
    ComboBoxObjName: TcxLookupComboBox;
    cxLabel3: TcxLabel;
    cxLabel1: TcxLabel;
    ADODataSet2: TADODataSet;
    DataSource2: TDataSource;
    SostMemo: TcxMemo;
    MemoPrim: TcxMemo;
    ButtonClearAll: TAdvGlowButton;
    procedure AdvGlowButton1Click(Sender: TObject);
    procedure ButtonOKClick(Sender: TObject);
    procedure ButtonClearAllClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmAddObjToExp: TFrmAddObjToExp;

implementation

uses
  Dialog2;

{$R *.dfm}

procedure TFrmAddObjToExp.AdvGlowButton1Click(Sender: TObject);
begin
  ObjName:='';
  close;
end;

procedure TFrmAddObjToExp.ButtonOKClick(Sender: TObject);
begin
  ObjName:=ComboBoxObjName.EditText;
  ObjSost:=SostMemo.Text;
  ObjPrim:=MemoPrim.Text;
  close;
end;

procedure TFrmAddObjToExp.ButtonClearAllClick(Sender: TObject);
begin
  SostMemo.Text:='';
  MemoPrim.Text:='';
  ComboBoxObjName.ItemIndex:=-1;
end;

end.
