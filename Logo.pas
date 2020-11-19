unit Logo;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, AdvPanel, AdvCircularProgress, rtflabel;

type
  TFrmLogo = class(TForm)
    ImageZnak: TImage;
    RTFLabel1: TRTFLabel;
    AdvCircularProgress1: TAdvCircularProgress;
    Timer1: TTimer;
    AdvPanel1: TAdvPanel;
    Label1: TLabel;
    Label2: TLabel;
    Timer2: TTimer;
    procedure Timer1Timer(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmLogo: TFrmLogo;

implementation

{$R *.dfm}

var
   p : integer;

procedure TFrmLogo.Timer1Timer(Sender: TObject);
begin
  close;
end;

procedure TFrmLogo.Timer2Timer(Sender: TObject);
begin
   p:=p+90;
   if p<100 then AdvCircularProgress1.Position:=p else Timer2.Enabled:=false;
end;

procedure TFrmLogo.FormCreate(Sender: TObject);
begin
   p:=1;
end;

end.
