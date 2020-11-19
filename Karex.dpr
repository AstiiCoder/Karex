program Karex;

uses
  Forms,
  MainWindow in 'MainWindow.pas' {FrmMainWindow},
  Grid1 in 'Grid1.pas' {FrmGrid1},
  Dialog1 in 'Dialog1.pas' {FrmDialog1},
  Dialog2 in 'Dialog2.pas' {FrmDialog2},
  AddObjToExp in 'AddObjToExp.pas' {FrmAddObjToExp},
  Period in 'Period.pas' {FrmPeriod},
  PeriodExp in 'PeriodExp.pas' {FrmPeriodExp},
  PeriodOtch in 'PeriodOtch.pas' {FrmPeriodOtch},
  About in 'About.pas' {FrmAbout},
  ProishObrOtech in 'ProishObrOtech.pas' {FrmProishObrOtech},
  AddObjToExpOtech in 'AddObjToExpOtech.pas' {FrmAddObjToExpOtech},
  ToolExcel in 'ToolExcel.pas' {FrmToolExcel},
  Graphik in 'Graphik.pas' {FrmGraph},
  ProishImpObr in 'ProishImpObr.pas' {FrmProishImpObr},
  Options in 'Options.pas' {FrmOptions},
  NumProtocols in 'NumProtocols.pas' {FrmNumProtocols},
  Logo in 'Logo.pas' {FrmLogo},
  Calendar in 'Calendar.pas' {FrmCalendar};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Карантинная экспертиза';
  Application.CreateForm(TFrmMainWindow, FrmMainWindow);
  Application.CreateForm(TFrmDialog2, FrmDialog2);
  Application.CreateForm(TFrmPeriod, FrmPeriod);
  Application.CreateForm(TFrmPeriodExp, FrmPeriodExp);
  Application.CreateForm(TFrmPeriodOtch, FrmPeriodOtch);
  Application.CreateForm(TFrmAddObjToExp, FrmAddObjToExp);
  Application.CreateForm(TFrmProishObrOtech, FrmProishObrOtech);
  Application.CreateForm(TFrmAddObjToExpOtech, FrmAddObjToExpOtech);
  Application.CreateForm(TFrmProishImpObr, FrmProishImpObr);
  //Application.CreateForm(TFrmGrid1, FrmGrid1);
  Application.Run;
end.
