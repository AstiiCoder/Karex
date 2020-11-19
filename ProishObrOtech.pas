unit ProishObrOtech;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxStyles, cxCustomData, cxGraphics, cxFilter, cxData,
  cxDataStorage, cxEdit, DB, cxDBData, cxLabel, cxContainer, cxTextEdit,
  cxGridLevel, cxClasses, cxControls, cxGridCustomView, 
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGrid,
  AdvGlowButton, ADODB, ImgList;

type
  TFrmProishObrOtech = class(TForm)
    ButtonOK: TAdvGlowButton;
    AdvGlowButton1: TAdvGlowButton;
    cxGrid1: TcxGrid;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    TextEditUlitsa: TcxTextEdit;
    cxLabel1: TcxLabel;
    cxLabel2: TcxLabel;
    ButtonAdd: TDBAdvGlowButton;
    ButtonEdit: TDBAdvGlowButton;
    ImageList1: TImageList;
    ADODataSet4: TADODataSet;
    DataSource4: TDataSource;
    procedure AdvGlowButton1Click(Sender: TObject);
    procedure ButtonOKClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmProishObrOtech: TFrmProishObrOtech;

implementation

uses
    MainWindow, Dialog1;

{$R *.dfm}

procedure TFrmProishObrOtech.AdvGlowButton1Click(Sender: TObject);
begin
  PerStart:='';
  close;
end;

procedure TFrmProishObrOtech.ButtonOKClick(Sender: TObject);
var
  s : string;
begin
  s:='';
  FrmDialog1.TextEditKodNasPunkt.Text:=FrmProishObrOtech.ADODataSet4.FieldByName('IDNASPUNKT').AsString;
  if  Length(FrmProishObrOtech.ADODataSet4.FieldByName('RAYONNAME').AsString)>1 then
    s:=FrmProishObrOtech.ADODataSet4.FieldByName('RAYONNAME').AsString;
  if  Length(FrmProishObrOtech.ADODataSet4.FieldByName('NASELPUNKTYPE').AsString)>1 then
    if s<>'' then s:=s+' '+FrmProishObrOtech.ADODataSet4.FieldByName('NASELPUNKTYPE').AsString;
  if  Length(FrmProishObrOtech.ADODataSet4.FieldByName('NASELPUNKTNAME').AsString)>1 then
    if s<>'' then s:=s+' '+FrmProishObrOtech.ADODataSet4.FieldByName('NASELPUNKTNAME').AsString;
  s:=FrmProishObrOtech.ADODataSet4.FieldByName('OBLNAME').AsString+' обл., '+s+' '+FrmProishObrOtech.TextEditUlitsa.Text;
  FrmDialog1.TextEditProishOtech.Text:=s;
  PerStart:=s;
  close;
end;

end.
