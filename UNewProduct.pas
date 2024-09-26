unit UNewProduct;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, acPNG, ExtCtrls, Buttons, StdCtrls, ADODB;

type
  TfrmNewProduct = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    edtCode: TEdit;
    Label2: TLabel;
    edtName: TEdit;
    Label3: TLabel;
    edtSpecs: TEdit;
    Label4: TLabel;
    edtStdCost: TEdit;
    Label5: TLabel;
    edtSellPrice: TEdit;
    Panel1: TPanel;
    Image1: TImage;
    lbMode: TLabel;
    Label7: TLabel;
    edtWhsStock: TEdit;
    Label8: TLabel;
    edtBrchStock: TEdit;
    btnSave: TBitBtn;
    procedure FormShow(Sender: TObject);
    procedure edtNameKeyPress(Sender: TObject; var Key: Char);
    procedure edtSpecsKeyPress(Sender: TObject; var Key: Char);
    procedure edtStdCostKeyPress(Sender: TObject; var Key: Char);
    procedure edtSellPriceKeyPress(Sender: TObject; var Key: Char);
    procedure edtWhsStockKeyPress(Sender: TObject; var Key: Char);
    procedure edtBrchStockKeyPress(Sender: TObject; var Key: Char);
    procedure btnProductSaveClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
  private
    { Private declarations }
    procedure insert_product(ID, vNAME, SPECIFICATION, STD_COST, LIST_PRICE, MIN_WHS_STOCK,
                              MIN_BCH_STOCK, DEL_FLAG : String);
    procedure update_product(ID, vNAME, SPECIFICATION, STD_COST, LIST_PRICE, MIN_WHS_STOCK,
                              MIN_BCH_STOCK, DEL_FLAG : String);
  public
    { Public declarations }
  end;

var
  frmNewProduct: TfrmNewProduct;

implementation

uses uMain, UMethods, UDataModule;

{$R *.dfm}

procedure TfrmNewProduct.FormShow(Sender: TObject);
begin
  if myProductMode = NewProduct then
    begin
      Self.Caption := 'New Product';
      lbMode.Caption := 'New Product';
      edtCode.Text := getNewPoductCode;
    end
  else if myProductMode = ProductUpdate then
    begin
      with frmMain.dbProducts do
        begin
          Self.Caption := 'Update Product';
          lbMode.Caption := 'Update Product';
          edtCode.Text := Fields[0].AsString;
          edtName.Text := Fields[1].AsString;
          edtSpecs.Text := Fields[2].AsString;
          edtStdCost.Text := Fields[3].AsString;
          edtSellPrice.Text := Fields[4].AsString;
          edtWhsStock.Text := Fields[5].AsString;
          edtBrchStock.Text := Fields[6].AsString;
        end;
    end;
  edtName.SetFocus;
end;



procedure TfrmNewProduct.edtNameKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    begin
      if Trim(edtName.Text) = '' then
        InputError(edtName, 2)
      else
        edtSpecs.SetFocus;
    end;
end;

procedure TfrmNewProduct.edtSpecsKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    begin
      if Trim(edtSpecs.Text) = '' then
        InputError(edtSpecs, 2)
      else
        edtStdCost.SetFocus;
    end;
end;

procedure TfrmNewProduct.edtStdCostKeyPress(Sender: TObject;
  var Key: Char);
var
  floatBuff : Double;
begin
  if not (Key in ['0' .. '9', #13, #8, '.']) then Key := #0;
  if Key = #13 then
    begin
      if trim(edtStdCost.Text) = '' then
        InputError(edtStdCost, 2)
      else if not tryStrToFloat(StringReplace(edtStdCost.Text, ',', '', [rfReplaceAll, rfIgnoreCase]), floatBuff) then
        InputError(edtStdCost, 2)
      else
        begin
           edtStdCost.Text := Copy(Format('%m', [floatBuff]), 2, Length(Format('%m', [floatBuff])));
           edtSellPrice.Setfocus;
        end;
    end;
end;

procedure TfrmNewProduct.edtSellPriceKeyPress(Sender: TObject;
  var Key: Char);
var
  floatBuff : Double;
begin
  if not (Key in ['0' .. '9', #13, #8, '.']) then Key := #0;
  if Key = #13 then
    begin
      if trim(edtSellPrice.Text) = '' then
        InputError(edtSellPrice, 2)
      else if not tryStrToFloat(StringReplace(edtSellPrice.Text, ',', '', [rfReplaceAll, rfIgnoreCase]), floatBuff) then
        InputError(edtSellPrice, 2)
      else
        begin
           edtSellPrice.Text := Copy(Format('%m', [floatBuff]), 2, Length(Format('%m', [floatBuff])));
           edtWhsStock.Setfocus;
        end;
    end;
end;

procedure TfrmNewProduct.edtWhsStockKeyPress(Sender: TObject;
  var Key: Char);
var
  intBuff : Integer;
begin
  if not (Key in ['0'..'9', #13, #8]) then Key :=#0;
  if Key = #13 then
    begin
      if trim(edtWhsStock.Text) = '' then
        InputError(edtWhsStock, 2)
      else if not TryStrtoInt(edtWhsStock.Text, intBuff) then
        InputError(edtWhsStock, 2)
      else
        edtBrchStock.SetFocus;
    end;
end;

procedure TfrmNewProduct.edtBrchStockKeyPress(Sender: TObject;
  var Key: Char);
var
  intBuff : Integer;
begin
  if not (Key in ['0'..'9', #13, #8]) then Key :=#0;
  if Key = #13 then
    begin
      if trim(edtBrchStock.Text) = '' then
        InputError(edtBrchStock, 2)
      else if not TryStrtoInt(edtBrchStock.Text, intBuff) then
        InputError(edtBrchStock, 2)
      else
        btnSave.SetFocus;
        
    end;
end;

procedure TfrmNewProduct.btnProductSaveClick(Sender: TObject);
var
  c : Char;
begin
  c := #13;
  FocusError := 0;
  if FocusError = 0 then edtName.OnKeyPress(Sender, c);
  if FocusError = 0 then edtSpecs.OnKeyPress(Sender, c);
  if FocusError = 0 then edtStdCost.OnKeyPress(Sender, c);
  if FocusError = 0 then edtSellPrice.OnKeyPress(Sender, c);
  if FocusError = 0 then edtWhsStock.OnKeyPress(Sender, c);
  if FocusError = 0 then edtBrchStock.OnKeyPress(Sender, c);
  if FocusError = 0 then
    begin
      if DisplayMessage(mtConfirmation, [mbYes, mbNo], 3)  = mrYes then
        try
          DM.adoConn.BeginTrans;
          if myProductMode = NewProduct then
            insert_product(Trim(edtCode.Text),
                            Trim(edtName.Text),
                            Trim(edtSpecs.Text),
                            StringReplace(Trim(edtStdCost.Text), ',', '', [rfReplaceAll, rfIgnoreCase]),
                            StringReplace(Trim(edtSellPrice.Text), ',', '', [rfReplaceAll, rfIgnoreCase]),
                            edtWhsStock.Text,
                            edtBrchStock.Text,
                            '0')
          else if myProductMode = ProductUpdate then
            update_product(Trim(edtCode.Text),
                            Trim(edtName.Text),
                            Trim(edtSpecs.Text),
                            StringReplace(Trim(edtStdCost.Text), ',', '', [rfReplaceAll, rfIgnoreCase]),
                            StringReplace(Trim(edtSellPrice.Text), ',', '', [rfReplaceAll, rfIgnoreCase]),
                            edtWhsStock.Text,
                            edtBrchStock.Text,
                            '0');
          DM.adoConn.CommitTrans;
          DisplayMessage(mtInformation, [mbOK], 4);
          ModalResult := mrOK;
        except
          DM.adoConn.RollbackTrans;
          DisplayMessage(mtError, [mbOK], 5);
        end;
    end;
end;

procedure TfrmNewProduct.insert_product(ID, vNAME, SPECIFICATION, STD_COST, LIST_PRICE, MIN_WHS_STOCK,
                              MIN_BCH_STOCK, DEL_FLAG : String);
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
      SQL.Add('INSERT INTO PRODUCTS(PRODUCT_ID, PRODUCT_NAME, SPECIFICATION, STD_COST, LIST_PRICE, MIN_WHS_STOCK, MIN_BCH_STOCK, DEL_FLAG)');
      SQL.Add('VALUES(:ID, :vNAME, :SPECIFICATION, :STD_COST, :LIST_PRICE, :MIN_WHS_STOCK, :MIN_BCH_STOCK, :DEL_FLAG)');
      Parameters.ParamByName('ID').Value := ID;
      Parameters.ParamByName('vNAME').Value := vNAME;
      Parameters.ParamByName('SPECIFICATION').Value := SPECIFICATION;
      Parameters.ParamByName('STD_COST').Value := STD_COST;
      Parameters.ParamByName('LIST_PRICE').Value := LIST_PRICE;
      Parameters.ParamByName('MIN_WHS_STOCK').Value := MIN_WHS_STOCK;
      Parameters.ParamByName('MIN_BCH_STOCK').Value := MIN_BCH_STOCK;
      Parameters.ParamByName('DEL_FLAG').Value := DEL_FLAG;
      if not Prepared then ExecSQL;
    end;
  FreeAndNil(qry);
end;


procedure TfrmNewProduct.update_product(ID, vNAME, SPECIFICATION, STD_COST,
  LIST_PRICE, MIN_WHS_STOCK, MIN_BCH_STOCK, DEL_FLAG: String);
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
      SQL.Add('UPDATE PRODUCTS');
      SQL.Add('set PRODUCT_NAME=:vNAME, SPECIFICATION=:SPECIFICATION, STD_COST=:STD_COST, LIST_PRICE=:LIST_PRICE, MIN_WHS_STOCK=:MIN_WHS_STOCK, MIN_BCH_STOCK=:MIN_BCH_STOCK, DEL_FLAG=:DEL_FLAG');
      SQL.Add('WHERE PRODUCT_ID=:ID');
      Parameters.ParamByName('ID').Value := ID;
      Parameters.ParamByName('vNAME').Value := vNAME;
      Parameters.ParamByName('SPECIFICATION').Value := SPECIFICATION;
      Parameters.ParamByName('STD_COST').Value := STD_COST;
      Parameters.ParamByName('LIST_PRICE').Value := LIST_PRICE;
      Parameters.ParamByName('MIN_WHS_STOCK').Value := MIN_WHS_STOCK;
      Parameters.ParamByName('MIN_BCH_STOCK').Value := MIN_BCH_STOCK;
      Parameters.ParamByName('DEL_FLAG').Value := DEL_FLAG;
      if not Prepared then ExecSQL;
    end;
  FreeAndNil(qry);
end;

procedure TfrmNewProduct.btnSaveClick(Sender: TObject);
var
  c : Char;
begin
  c := #13;
  FocusError := 0;
  if FocusError = 0 then edtName.OnKeyPress(Sender, c);
  if FocusError = 0 then edtSpecs.OnKeyPress(Sender, c);
  if FocusError = 0 then edtStdCost.OnKeyPress(Sender, c);
  if FocusError = 0 then edtSellPrice.OnKeyPress(Sender, c);
  if FocusError = 0 then edtWhsStock.OnKeyPress(Sender, c);
  if FocusError = 0 then edtBrchStock.OnKeyPress(Sender, c);
  if FocusError = 0 then
    begin
      if DisplayMessage(mtConfirmation, [mbYes, mbNo], 3)  = mrYes then
        try
          DM.adoConn.BeginTrans;
          if myProductMode = NewProduct then
            insert_product(Trim(edtCode.Text),
                            Trim(edtName.Text),
                            Trim(edtSpecs.Text),
                            StringReplace(Trim(edtStdCost.Text), ',', '', [rfReplaceAll, rfIgnoreCase]),
                            StringReplace(Trim(edtSellPrice.Text), ',', '', [rfReplaceAll, rfIgnoreCase]),
                            edtWhsStock.Text,
                            edtBrchStock.Text,
                            '0')
          else if myProductMode = ProductUpdate then
            update_product(Trim(edtCode.Text),
                            Trim(edtName.Text),
                            Trim(edtSpecs.Text),
                            StringReplace(Trim(edtStdCost.Text), ',', '', [rfReplaceAll, rfIgnoreCase]),
                            StringReplace(Trim(edtSellPrice.Text), ',', '', [rfReplaceAll, rfIgnoreCase]),
                            edtWhsStock.Text,
                            edtBrchStock.Text,
                            '0');
          DM.adoConn.CommitTrans;
          DisplayMessage(mtInformation, [mbOK], 4);
          ModalResult := mrOK;
        except
          DM.adoConn.RollbackTrans;
          DisplayMessage(mtError, [mbOK], 5);
        end;
    end;
end;
end.
