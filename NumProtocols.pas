unit NumProtocols;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, cxStyles,
  cxCustomData, cxGraphics, cxFilter, cxData,
  cxDataStorage, cxEdit, cxDBData, cxGridLevel, cxClasses, cxControls,
  cxGridCustomView, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxGrid;

type
  TFrmNumProtocols = class(TForm)
    ADODataSet2: TADODataSet;
    DataSource2: TDataSource;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    cxGrid1: TcxGrid;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmNumProtocols: TFrmNumProtocols;

implementation

uses
  MainWindow;

{$R *.dfm}

end.
