unit Graphik;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxStyles, cxCustomData, cxGraphics, cxFilter, cxData,
  cxDataStorage, cxEdit, DB, cxDBData, AdvGlowButton, AdvToolBar,
  cxGridLevel, cxClasses, cxControls, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGrid,
  cxGridChartView, cxGridDBChartView, ADODB, cxLabel, cxContainer,
  cxTextEdit, cxMaskEdit, cxDropDownEdit;

type
  TFrmGraph = class(TForm)
    cxGrid1: TcxGrid;
    cxGrid1Level1: TcxGridLevel;
    AdvToolBarPager1: TAdvToolBarPager;
    AdvPage1: TAdvPage;
    AdvToolBar1: TAdvToolBar;
    ButtonSave: TAdvGlowButton;
    cxGrid1DBChartView1: TcxGridDBChartView;
    ADODataSet2: TADODataSet;
    DataSource2: TDataSource;
    ButtonCreateGraph: TAdvGlowButton;
    ComboBoxOsi: TcxComboBox;
    ComboBoxVal: TcxComboBox;
    cxLabel1: TcxLabel;
    cxLabel2: TcxLabel;
    AdvToolBarSeparator1: TAdvToolBarSeparator;
    AdvGlowButton1: TAdvGlowButton;
    procedure ButtonCreateGraphClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmGraph: TFrmGraph;

implementation

uses
  MainWindow, Grid1;

{$R *.dfm}

procedure TFrmGraph.ButtonCreateGraphClick(Sender: TObject);
var
  ASeriesItog: TcxGridDBChartSeries;
  i: integer;
  Os, Va: string;
begin
  for i:=0 to FrmGrid1.cxGrid1DBTableView1.ColumnCount-1 do
    begin
      if FrmGrid1.cxGrid1DBTableView1.Columns[i].Caption=ComboBoxOsi.SelText then Os:=FrmGrid1.cxGrid1DBTableView1.Columns[i].DataBinding.FieldName;
      if FrmGrid1.cxGrid1DBTableView1.Columns[i].Caption=ComboBoxVal.SelText then Va:=FrmGrid1.cxGrid1DBTableView1.Columns[i].DataBinding.FieldName;
    end;
  ADODataSet2.Active:=false;
  ADODataSet2.CommandText:=FrmMainWindow.ADODataSet1.CommandText;
  ADODataSet2.Active:=true;
  cxGrid1DBChartView1.ClearSeries;
  ASeriesItog := cxGrid1DBChartView1.CreateSeries;
  ASeriesItog.DataBinding.FieldName := Va;
  cxGrid1DBChartView1.Categories.DataBinding.FieldName := Os;
  ASeriesItog.DisplayText:='сумма';
end;

end.
