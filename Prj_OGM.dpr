program Prj_OGM;

uses
  Forms,
  uMain in 'uMain.pas' {frmMain},
  UNewProduct in 'UNewProduct.pas' {frmNewProduct},
  UMethods in 'UMethods.pas',
  UDataModule in 'UDataModule.pas' {DM: TDataModule};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Ongpeng General Merchandise Inventory System';
  Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TfrmNewProduct, frmNewProduct);
  Application.CreateForm(TDM, DM);
  Application.Run;
end.
