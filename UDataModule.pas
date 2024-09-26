unit UDataModule;

interface

uses
  SysUtils, Classes, DB, ADODB, Provider, DBClient;

type
  TDM = class(TDataModule)
    adoConn: TADOConnection;
    dtsProduct: TDataSource;
    qryProducts: TADOQuery;
    qrySelectProduct: TADOQuery;
    qryOrders: TADOQuery;
    dtsOrders: TDataSource;
    qryRelease: TADOQuery;
    dtsRelease: TDataSource;
    qrySales: TADOQuery;
    dtsSales: TDataSource;
    qryReport: TADOQuery;
    dtsReport: TDataSource;
    qryCDSProducts: TADOQuery;
    dtsCDSProducts: TDataSource;
    cdsWarehouseInv: TClientDataSet;
    dspWarehouseInv: TDataSetProvider;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DM: TDM;

implementation

{$R *.dfm}

end.
