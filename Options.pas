unit Options;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxControls, cxContainer, cxEdit, cxCheckBox, AdvGlowButton, Registry;

type
  TFrmOptions = class(TForm)
    ButtonOK: TAdvGlowButton;
    AdvGlowButton1: TAdvGlowButton;
    CheckBox_Oeape: TcxCheckBox;
    CheckBox_MultiUserReg: TcxCheckBox;
    CheckBox_ViewStatistics: TcxCheckBox;
    procedure AdvGlowButton1Click(Sender: TObject);
    procedure ButtonOKClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmOptions: TFrmOptions;

implementation

uses
  MainWindow;

{$R *.dfm}

procedure TFrmOptions.AdvGlowButton1Click(Sender: TObject);
begin
   close;
end;

procedure TFrmOptions.ButtonOKClick(Sender: TObject);
var
   REG : TRegistry;
begin
   //новые значнения ключей
   REG := TRegistry.Create;
   REG.RootKey:=HKEY_CURRENT_USER;
   REG.OpenKey('Software\Karex\Local\Settings',true);
   REG.WriteBool('OpenExcelAfterProtEnt',CheckBox_Oeape.Checked);
   REG.WriteBool('MultiUserReg',CheckBox_MultiUserReg.Checked);
   REG.WriteBool('ViewStatistics',CheckBox_ViewStatistics.Checked);
   //REG.WriteInteger('Top',Top);
   REG.CloseKey;
   REG.Destroy;
   //новые значения переменных
   Oeape:=CheckBox_Oeape.Checked;
   MUReg:=CheckBox_MultiUserReg.Checked;
   ViewStatistic:=CheckBox_ViewStatistics.Checked;
   FrmMainWindow.WiiProgressBar1.Visible:=ViewStatistic;
   FrmMainWindow.Timer2.Enabled:=true;
   close;
end;

procedure TFrmOptions.FormShow(Sender: TObject);
begin
   //восстановим значения флажков
   CheckBox_Oeape.Checked:=Oeape;
   CheckBox_MultiUserReg.Checked:=MUReg;
   CheckBox_ViewStatistics.Checked:=ViewStatistic;
end;

end.
