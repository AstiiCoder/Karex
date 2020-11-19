unit ToolExcel;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, AdvGlowButton, AdvToolBar;

type
  TFrmToolExcel = class(TForm)
    ToolBarExcelFonts: TAdvToolBar;
    ButtonSmallFont: TAdvGlowButton;
    ButtonLargeFont: TAdvGlowButton;
    ButtonShowExcel: TAdvGlowButton;
    ButtonDrawBordert: TAdvGlowButton;
    ButtonClearBordert: TAdvGlowButton;
    procedure ButtonLargeFontClick(Sender: TObject);
    procedure ButtonSmallFontClick(Sender: TObject);
    procedure ButtonDrawBordertClick(Sender: TObject);
    procedure ButtonClearBordertClick(Sender: TObject);
    procedure ButtonShowExcelClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmToolExcel: TFrmToolExcel;

implementation

uses
  Grid1, Dialog1;

{$R *.dfm}

procedure TFrmToolExcel.ButtonLargeFontClick(Sender: TObject);
begin
  FrmGrid1.ButtonLargeFontClick(Sender);
end;

procedure TFrmToolExcel.ButtonSmallFontClick(Sender: TObject);
begin
  FrmGrid1.ButtonSmallFontClick(Sender);
end;

procedure TFrmToolExcel.ButtonDrawBordertClick(Sender: TObject);
begin
  FrmGrid1.ButtonDrawBordertClick(Sender);
end;

procedure TFrmToolExcel.ButtonClearBordertClick(Sender: TObject);
begin
  FrmGrid1.ButtonClearBordertClick(Sender);
end;

procedure TFrmToolExcel.ButtonShowExcelClick(Sender: TObject);
begin
  FrmGrid1.ButtonShowExcelClick(Sender);
end;

procedure TFrmToolExcel.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  try
    if FrmDialog1.Caption<>'' then FrmDialog1.ButtonExcelTool.Visible:=true;
  except
  end;
end;

end.
