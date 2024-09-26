unit UMethods;

interface

uses Dialogs, Graphics, DB, SysUtils, StdCtrls, Forms, Types, Windows,
    DBTables, Classes, ADODB, DBClient, ComCtrls;

  function MessageDlg(const Msg: String; DlgType: TMsgDlgType;
                      Buttons: TMsgDlgButtons; HelpCtx: Integer): Integer;
  function DisplayMessage(AType: TMsgDlgType; AButtons: TMsgDlgButtons;
                      Index: Integer): Word;
  procedure SelectError(var ComboBox: TComboBox; Index: Integer);
  procedure InputError(var Edit: TEdit; Index: Integer);
  function getConnString() : String;
  procedure dbConnect(connString : String);
  procedure dbDisconnect();
  procedure displayProducts(ASearch : String);  overload;
  procedure displayProducts();  overload;
  function getNewPoductCode : String;
  function getNewOrderCode() : String;
  function getNewRelCode() : String;
  function getNewSalesCode() : String;
  procedure insert_orders(ORDER_NO, ORDER_DATE, ORDER_PRODUCT_ID, ORDER_QUANTITY, ORDER_DEL_FLAG : String);
  procedure insert_release(RELEASE_NO, RELEASE_DATE, RELEASE_PRODUCT_ID, RELEASE_QUANTITY, RELEASE_BRANCH_CODE, RELEASE_DEL_FLAG : String);
  procedure insert_sales(SALES_NO, SALES_DATE, SALES_PRODUCT_ID, SALES_QUANTITY, SALES_SALE_PRICE, SALES_BRANCH_CODE, SALES_DEL_FLAG : String);
  procedure getCDSProducts(ASearch : String);

type
  TProductMode = (NewProduct, ProductUpdate);


var
  FocusError : Integer;
  BranchCode : String;
  myProductMode : TProductMode;
  InputChecker : Boolean;
  cdsWarehouseInv : TClientDataSet;

implementation

uses UDataModule;

function MessageDlg(const Msg: String; DlgType: TMsgDlgType;
  Buttons: TMsgDlgButtons; HelpCtx: Integer): Integer;
begin
  with CreateMessageDialog(Msg, DlgType, Buttons) do
    try
      Position := poOwnerFormCenter;
      Result := ShowModal;
    finally
      Free;
    end;
end;

function DisplayMessage(AType: TMsgDlgType; AButtons: TMsgDlgButtons;
  Index: Integer): Word;
var
  s: String;
begin
  case Index of
  1: s := 'Database connection failed. Please check your database configuration file!';
  2: s := 'Please enter required or correct values for this field.';
  3: s := 'Are you sure you want to proceed with your transaction';
  4: s := 'Transaction completed.';
  5: s := 'An error occured during the process of your transaction, Please report to your system developer!';
  6: s := 'Product information missing. Please check your product entries!';
  7: s := 'You do not have enough stocks for this product.';
  end;
  DisplayMessage := MessageDlg(s, AType, AButtons, 0);
end;

procedure InputError(var Edit: TEdit; Index: Integer);
begin
  DisplayMessage(mtError, [mbOK], Index);
  Edit.SelectAll;
  Edit.SetFocus;
  FocusError := 1;
end;

procedure SelectError(var ComboBox: TComboBox; Index: Integer);
begin
  DisplayMessage(mtError, [mbOK], Index);
  ComboBox.SelectAll;
  ComboBox.SetFocus;
  FocusError := 1;
end;

function getNewPoductCode: String;
var
  qry : TADOQuery;
begin
  qry :=  TADOQuery.Create(nil);
  with qry do
    begin
      Connection := DM.adoConn;
      //dbDisconnect;
      dbConnect(getConnString());
      SQL.Add('select max(PRODUCT_ID) + 1 from products');
      if not Prepared then Open;
      if IsEmpty then
        Result := '00000000'
      else
        Result :=  Format('%.8d', [Fields[0].AsInteger]);
    end;
  FreeAndNil(qry);
end;

function getNewOrderCode() : String;
var
  qry : TADOQuery;
begin
  qry :=  TADOQuery.Create(nil);
  with qry do
    begin
      Connection := DM.adoConn;
      //dbDisconnect;
      dbConnect(getConnString());
      SQL.Add('select max(ORDER_NO) + 1 from orders');
      if not Prepared then Open;
      if IsEmpty then
        Result := '00000000'
      else
        Result :=  Format('%.8d', [Fields[0].AsInteger]);
    end;
  FreeAndNil(qry);
end;

function getNewRelCode() : String;
var
  qry : TADOQuery;
begin
  qry :=  TADOQuery.Create(nil);
  with qry do
    begin
      Connection := DM.adoConn;
      //dbDisconnect;
      dbConnect(getConnString());
      SQL.Add('select max(RELEASE_NO) + 1 from release');
      if not Prepared then Open;
      if IsEmpty then
        Result := '00000000'
      else
        Result :=  Format('%.8d', [Fields[0].AsInteger]);
    end;
  FreeAndNil(qry);
end;


function getNewSalesCode() : String;
var
  qry : TADOQuery;
begin
  qry :=  TADOQuery.Create(nil);
  with qry do
    begin
      Connection := DM.adoConn;
      //dbDisconnect;
      dbConnect(getConnString());
      SQL.Add('select max(SALES_NO) + 1 from sales');
      if not Prepared then Open;
      if IsEmpty then
        Result := '00000000'
      else
        Result :=  Format('%.8d', [Fields[0].AsInteger]);
    end;
  FreeAndNil(qry);
end;

procedure dbDisconnect;
begin
  if DM.adoConn.Connected then
    DM.adoConn.Close;
end;

procedure displayProducts;
begin
  with DM.qryProducts do
    begin
      Connection := DM.adoConn;
      //dbDisconnect;
      dbConnect(getConnString());
      SQL.Clear;
      SQL.Add('select PRODUCT_ID, PRODUCT_NAME, SPECIFICATION,');
      SQL.Add('format(STD_COST, "Standard") as STD_COST,');
      SQL.Add('format(LIST_PRICE, "Standard") as LIST_PRICE,');
      SQL.Add('MIN_WHS_STOCK, MIN_BCH_STOCK from products');
      SQL.Add('where DEL_FLAG = ''0''');
      if not Prepared then Open;
    end;
end;

procedure dbConnect(connString : String);
begin
  if not DM.adoConn.Connected = True then
    begin
      DM.adoConn.ConnectionString := connString;
      DM.adoConn.Connected;
    end;
end;

procedure displayProducts(ASearch : String);
var
  vSearch : String;
begin
  vSearch :=  '%' +  ASearch +  '%';
  with DM.qryProducts do
    begin
      Connection := DM.adoConn;
      //dbDisconnect;
      dbConnect(getConnString());
      SQL.Clear;
      SQL.Add('select PRODUCT_ID, PRODUCT_NAME, SPECIFICATION,');
      SQL.Add('format(STD_COST, "Standard") as STD_COST,');
      SQL.Add('format(LIST_PRICE, "Standard") as LIST_PRICE,');
      SQL.Add('MIN_WHS_STOCK, MIN_BCH_STOCK from products');
      SQL.Add('where ((PRODUCT_NAME like '+ QuotedStr(vSearch) + ') or (SPECIFICATION like ' + QuotedStr(vSearch) + '))' );
      SQL.Add('and DEL_FLAG = ''0''');
      if not Prepared then Open;
    end;
end;

procedure getCDSProducts(ASearch : String);
var
  vSearch : String;
begin
  vSearch :=  '%' +  ASearch +  '%';
  with DM.qryCDSProducts do
    begin
      Connection := DM.adoConn;
      //dbDisconnect;
      dbConnect(getConnString());
      SQL.Clear;
      SQL.Add('select PRODUCT_ID, PRODUCT_NAME, SPECIFICATION,');
      SQL.Add('format(STD_COST, "Standard") as STD_COST,');
      SQL.Add('format(LIST_PRICE, "Standard") as LIST_PRICE,');
      SQL.Add('MIN_WHS_STOCK, MIN_BCH_STOCK from products');
      SQL.Add('where ((PRODUCT_NAME like '+ QuotedStr(vSearch) + ') or (SPECIFICATION like ' + QuotedStr(vSearch) + '))' );
      SQL.Add('and DEL_FLAG = ''0''');
      if not Prepared then Open;
    end;
end;


function getConnString: String;
var
  F : TextFile;
  s : string;
begin
  if FileExists(ExtractFilePath(Application.ExeName) + '\connString.ini') then
    begin
      try
        AssignFile(F ,ExtractFilePath(Application.ExeName) + '\connString.ini');
        Reset(F);
        Readln(F,s);
        Result := s;
      finally
        CloseFile(F)
      end
    end
  else DisplayMessage(mtError, [mbOK], 1);
end;

procedure insert_orders(ORDER_NO, ORDER_DATE, ORDER_PRODUCT_ID, ORDER_QUANTITY, ORDER_DEL_FLAG : String);
var
  qry : TADOQuery;
begin
  qry := TADOQuery.Create(nil);
  with qry do
    begin
      Connection := DM.adoConn;
      //dbDisconnect;
      dbConnect(getConnString());
      SQL.Clear;
      SQL.Add('INSERT INTO ORDERS(ORDER_NO, ORDER_DATE, ORDER_PRODUCT_ID, ORDER_QUANTITY, ORDER_DEL_FLAG)');
      SQL.Add('VALUES(:ORDER_NO, :ORDER_DATE, :ORDER_PRODUCT_ID, :ORDER_QUANTITY, :ORDER_DEL_FLAG)');
      Parameters.ParamByName('ORDER_NO').Value := ORDER_NO;
      Parameters.ParamByName('ORDER_DATE').Value := ORDER_DATE;
      Parameters.ParamByName('ORDER_PRODUCT_ID').Value := ORDER_PRODUCT_ID;
      Parameters.ParamByName('ORDER_QUANTITY').Value := ORDER_QUANTITY;
      Parameters.ParamByName('ORDER_DEL_FLAG').Value := ORDER_DEL_FLAG;
      if not Prepared then ExecSQL;
    end;
  FreeAndNil(qry);
end;

procedure insert_release(RELEASE_NO, RELEASE_DATE, RELEASE_PRODUCT_ID, RELEASE_QUANTITY, RELEASE_BRANCH_CODE, RELEASE_DEL_FLAG : String);
var
  qry : TADOQuery;
begin
  qry := TADOQuery.Create(nil);
  with qry do
    begin
      Connection := DM.adoConn;
      //dbDisconnect;
      dbConnect(getConnString());
      SQL.Clear;
      SQL.Add('INSERT INTO RELEASE(RELEASE_NO, RELEASE_DATE, RELEASE_PRODUCT_ID, RELEASE_QUANTITY, RELEASE_BRANCH_CODE, RELEASE_DEL_FLAG)');
      SQL.Add('VALUES(:RELEASE_NO, :RELEASE_DATE, :RELEASE_PRODUCT_ID, :RELEASE_QUANTITY, :RELEASE_BRANCH_CODE, :RELEASE_DEL_FLAG)');
      Parameters.ParamByName('RELEASE_NO').Value := RELEASE_NO;
      Parameters.ParamByName('RELEASE_DATE').Value := RELEASE_DATE;
      Parameters.ParamByName('RELEASE_PRODUCT_ID').Value := RELEASE_PRODUCT_ID;
      Parameters.ParamByName('RELEASE_QUANTITY').Value := RELEASE_QUANTITY;
      Parameters.ParamByName('RELEASE_BRANCH_CODE').Value := RELEASE_BRANCH_CODE;
      Parameters.ParamByName('RELEASE_DEL_FLAG').Value := RELEASE_DEL_FLAG;
      if not Prepared then ExecSQL;
    end;
  FreeAndNil(qry);
end;

procedure insert_sales(SALES_NO, SALES_DATE, SALES_PRODUCT_ID, SALES_QUANTITY, SALES_SALE_PRICE, SALES_BRANCH_CODE, SALES_DEL_FLAG : String);
var
  qry : TADOQuery;
begin
  qry := TADOQuery.Create(nil);
  with qry do
    begin
      Connection := DM.adoConn;
      //dbDisconnect;
      dbConnect(getConnString());
      SQL.Clear;
      SQL.Add('INSERT INTO SALES(SALES_NO, SALES_DATE, SALES_PRODUCT_ID, SALES_QUANTITY, SALES_SALE_PRICE, SALES_BRANCH_CODE, SALES_DEL_FLAG)');
      SQL.Add('VALUES(:SALES_NO, :SALES_DATE, :SALES_PRODUCT_ID, :SALES_QUANTITY, :SALES_SALE_PRICE, :SALES_BRANCH_CODE, :SALES_DEL_FLAG)');
      Parameters.ParamByName('SALES_NO').Value := SALES_NO;
      Parameters.ParamByName('SALES_DATE').Value := SALES_DATE;
      Parameters.ParamByName('SALES_PRODUCT_ID').Value := SALES_PRODUCT_ID;
      Parameters.ParamByName('SALES_QUANTITY').Value := SALES_QUANTITY;
      Parameters.ParamByName('SALES_SALE_PRICE').Value := SALES_SALE_PRICE;
      Parameters.ParamByName('SALES_BRANCH_CODE').Value := SALES_BRANCH_CODE;
      Parameters.ParamByName('SALES_DEL_FLAG').Value := SALES_DEL_FLAG;
      if not Prepared then ExecSQL;
    end;
  FreeAndNil(qry);
end;

end.
