unit About;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, AdvGlowButton, cxControls, cxContainer, cxEdit, cxLabel,
  ExtCtrls, cxTextEdit;

type
  TFrmAbout = class(TForm)
    Image1: TImage;
    cxLabel1: TcxLabel;
    ButtonOK: TAdvGlowButton;
    cxLabel2: TcxLabel;
    cxLabel3: TcxLabel;
    cxLabel4: TcxLabel;
    cxLabel5: TcxLabel;
    cxTextEdit1: TcxTextEdit;
    procedure ButtonOKClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmAbout: TFrmAbout;

implementation

{$R *.dfm}

procedure TFrmAbout.ButtonOKClick(Sender: TObject);
begin
  close;
end;

end.
