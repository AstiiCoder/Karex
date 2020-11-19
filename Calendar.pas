unit Calendar;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, AdvSmoothCalendar;

type
  TFrmCalendar = class(TForm)
    AdvSmoothCalendar1: TAdvSmoothCalendar;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmCalendar: TFrmCalendar;

implementation

{$R *.dfm}

end.
