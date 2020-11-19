unit ProishImpObr;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxControls, cxContainer, cxListBox, AdvGlowButton, DB, ADODB,
  cxDBEdit, cxStyles, cxCustomData, cxGraphics, cxFilter, cxData,
  cxDataStorage, cxEdit, cxDBData, cxGridLevel, cxClasses,
  cxGridCustomView, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxGrid;

type
  TFrmProishImpObr = class(TForm)
    ButtonOK: TAdvGlowButton;
    AdvGlowButton1: TAdvGlowButton;
    ADODataSet2: TADODataSet;
    DataSource1: TDataSource;
    cxGrid1: TcxGrid;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    ADOCommand1: TADOCommand;
    procedure AdvGlowButton1Click(Sender: TObject);
    procedure ButtonOKClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmProishImpObr: TFrmProishImpObr;

implementation

uses
  MainWindow, Grid1;

{$R *.dfm}

procedure TFrmProishImpObr.AdvGlowButton1Click(Sender: TObject);
begin
  PerStart:='';
  ADODataSet2.Active:=false;
  cxGrid1DBTableView1.ClearItems;
  close;
end;

procedure TFrmProishImpObr.ButtonOKClick(Sender: TObject);
begin
  PerStart:=cxGrid1DBTableView1.DataController.DataSource.DataSet.FieldByName('IDPROISH').AsString;
  ADODataSet2.Active:=false;
  cxGrid1DBTableView1.ClearItems;
  close;
end;

end.
