program POS;

uses
  Forms,
  UMain in 'UMain.pas' {FrMain},
  UCommon in 'UCommon.pas',
  UModulTemp in 'UModulTemp.pas' {dmTemp: TDataModule};

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'POS Monitoring by SAHA';
  Application.CreateForm(TFrMain, FrMain);
  Application.CreateForm(TdmTemp, dmTemp);
  Application.Run;
end.
