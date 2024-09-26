unit uMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, Buttons, Grids, DBGrids, ExtCtrls, DB, DBClient, ADODB, StdCtrls,
  Menus, TaskDialog, TaskDialogEx, scExcelExport, abcbusy;

type
  TfrmMain = class(TForm)
    pageControl: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet5: TTabSheet;
    Panel1: TPanel;
    dbProducts: TDBGrid;
    edtSearchProduct: TEdit;
    Label1: TLabel;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    Panel2: TPanel;
    pmProduct: TPopupMenu;
    NewProduct1: TMenuItem;
    UpdateProduct1: TMenuItem;
    DeleteProduct1: TMenuItem;
    Refresh1: TMenuItem;
    GroupBox1: TGroupBox;
    Label3: TLabel;
    Label4: TLabel;
    Label2: TLabel;
    dtpOrderFrom: TDateTimePicker;
    dtpOrderTo: TDateTimePicker;
    edtOrderSearch: TEdit;
    GroupBox2: TGroupBox;
    dbOrders: TDBGrid;
    lvOrders: TListView;
    Label7: TLabel;
    dtpDelDate: TDateTimePicker;
    Label5: TLabel;
    cbOrderProduct: TComboBox;
    Label6: TLabel;
    edtOrderQty: TEdit;
    Label8: TLabel;
    edtOrderName: TEdit;
    Label9: TLabel;
    edtOrderSpecs: TEdit;
    Label10: TLabel;
    edtOrderStdCost: TEdit;
    Label11: TLabel;
    edtOrderSellPrice: TEdit;
    Label12: TLabel;
    edtOrderCode: TEdit;
    btnOrdersAdd: TBitBtn;
    btnOrdersPost: TBitBtn;
    btnProductSearch: TBitBtn;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    BitBtn3: TBitBtn;
    BitBtn4: TBitBtn;
    pmOrderList: TPopupMenu;
    Remove: TMenuItem;
    ClearAll: TMenuItem;
    pmOrders: TPopupMenu;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    Panel3: TPanel;
    GroupBox3: TGroupBox;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label20: TLabel;
    lvRelease: TListView;
    dtpRelDate: TDateTimePicker;
    cbRelProduct: TComboBox;
    edtRelQty: TEdit;
    edtRelName: TEdit;
    edtRelSpecs: TEdit;
    edtRelInventory: TEdit;
    edtRelCode: TEdit;
    btnRelAdd: TBitBtn;
    btnRelPost: TBitBtn;
    GroupBox4: TGroupBox;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    dtpRelFrom: TDateTimePicker;
    dtpRelTo: TDateTimePicker;
    edtRelSearch: TEdit;
    BitBtn7: TBitBtn;
    dbRelease: TDBGrid;
    pmRelList: TPopupMenu;
    Remove1: TMenuItem;
    ClearAll1: TMenuItem;
    ibx_branch: TAdvInputTaskDialogEx;
    pmRelease: TPopupMenu;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    Panel4: TPanel;
    GroupBox5: TGroupBox;
    Label19: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    Label27: TLabel;
    Label28: TLabel;
    Label29: TLabel;
    lvSell: TListView;
    dtpSellDate: TDateTimePicker;
    cbSellProduct: TComboBox;
    edtSellQty: TEdit;
    edtSellName: TEdit;
    EdtSellSpecs: TEdit;
    edtBranchInventory: TEdit;
    edtSellCode: TEdit;
    btnSellAdd: TBitBtn;
    BitBtn6: TBitBtn;
    GroupBox6: TGroupBox;
    Label30: TLabel;
    Label31: TLabel;
    Label32: TLabel;
    dtpSellFrom: TDateTimePicker;
    dtpSellTo: TDateTimePicker;
    edtSellSearch: TEdit;
    btnSellSearch: TBitBtn;
    dbSell: TDBGrid;
    Label33: TLabel;
    cbBranch: TComboBox;
    Label34: TLabel;
    edtSellPrice: TEdit;
    pmSellList: TPopupMenu;
    Remove2: TMenuItem;
    ClearAll2: TMenuItem;
    pmSales: TPopupMenu;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    btnReport: TBitBtn;
    Label36: TLabel;
    cbReportType: TComboBox;
    dbReport: TDBGrid;
    edtReportSearch: TEdit;
    Label35: TLabel;
    scExcelExport: TscExcelExport;
    Label37: TLabel;
    BitBtn8: TBitBtn;
    Label38: TLabel;
    aBsy_main: TabcBusy;
    ExportOrders: TMenuItem;
    ExportToExcel1: TMenuItem;
    ExportToExcel2: TMenuItem;
    ExportToExcel3: TMenuItem;
    procedure SpeedButton1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure edtSearchProductKeyPress(Sender: TObject; var Key: Char);
    procedure btnSeachProductClick(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure pageControlChange(Sender: TObject);
    procedure cbOrderProductKeyPress(Sender: TObject; var Key: Char);
    procedure edtOrderQtyKeyPress(Sender: TObject; var Key: Char);
    procedure btnProductSearchClick(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure btnOrdersAddClick(Sender: TObject);
    procedure pmOrderListPopup(Sender: TObject);
    procedure RemoveClick(Sender: TObject);
    procedure ClearAllClick(Sender: TObject);
    procedure btnOrdersPostClick(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure MenuItem2Click(Sender: TObject);
    procedure cbRelProductKeyPress(Sender: TObject; var Key: Char);
    procedure edtRelQtyKeyPress(Sender: TObject; var Key: Char);
    procedure btnRelAddClick(Sender: TObject);
    procedure Remove1Click(Sender: TObject);
    procedure ClearAll1Click(Sender: TObject);
    procedure pmRelListPopup(Sender: TObject);
    procedure btnRelPostClick(Sender: TObject);
    procedure BitBtn7Click(Sender: TObject);
    procedure MenuItem3Click(Sender: TObject);
    procedure MenuItem4Click(Sender: TObject);
    procedure cbBranchChange(Sender: TObject);
    procedure cbSellProductKeyPress(Sender: TObject; var Key: Char);
    procedure edtSellQtyKeyPress(Sender: TObject; var Key: Char);
    procedure edtSellPriceKeyPress(Sender: TObject; var Key: Char);
    procedure btnSellAddClick(Sender: TObject);
    procedure BitBtn6Click(Sender: TObject);
    procedure btnSellSearchClick(Sender: TObject);
    procedure ClearAll2Click(Sender: TObject);
    procedure Remove2Click(Sender: TObject);
    procedure pmSellListPopup(Sender: TObject);
    procedure MenuItem5Click(Sender: TObject);
    procedure MenuItem6Click(Sender: TObject);
    procedure edtOrderSearchKeyPress(Sender: TObject; var Key: Char);
    procedure edtRelSearchKeyPress(Sender: TObject; var Key: Char);
    procedure edtSellSearchKeyPress(Sender: TObject; var Key: Char);
    procedure btnReportClick(Sender: TObject);
    procedure BitBtn8Click(Sender: TObject);
    procedure ExportOrdersClick(Sender: TObject);
    procedure ExportToExcel1Click(Sender: TObject);
    procedure ExportToExcel2Click(Sender: TObject);
    procedure ExportToExcel3Click(Sender: TObject);
  private
    { Private declarations }
    procedure delete_product(ID : String);
    procedure delete_orders(ID : String);
    procedure delete_releases(ID: String);
    procedure delete_sales(ID: String);
    function getProducts : TStrings;
    procedure displayOrders(aSearch : String);
    procedure displayReleases(aSearch : String);
    procedure displaySales(aSearch: String);
  public
    { Public declarations }
    procedure initProductpage;
    procedure initOrderspage;
    procedure initReleasepage;
    procedure initSalespage;


    procedure getWarehouseInventory(ASearch : String);
    procedure getBranchInventory(branchCode, ASearch : String);
    procedure getSales(ASearch : String);
    function check_products(NameSpecs : String) :  Boolean;
    function getSumOrders(ProductID : String) : Integer;
    function getSumRelease(ProductID : String) : Integer;
    function getSumBranchRelease(ProductID, BranchID : String) : Integer;
    function getSumBranchSalesProduct(ProductID, BranchID : String) : Integer;
    procedure FitGrid(Grid: TDBGrid);
    procedure SetDBGridColumn(Grid: TDBGrid);
end;


var
  frmMain: TfrmMain;

implementation

uses UNewProduct, UMethods, UDataModule;

{$R *.dfm}

{ TfrmMain }


procedure TfrmMain.SpeedButton1Click(Sender: TObject);
begin
  frmNewProduct :=  TfrmNewProduct.Create(Self);
    with frmNewProduct do
      begin
        try
          Hide;
          myProductMode := NewProduct;
          ShowModal;
        finally
          Free;
          displayProducts();
        end;
      end
end;

procedure TfrmMain.FormShow(Sender: TObject);
begin
  displayProducts();
end;

procedure TfrmMain.edtSearchProductKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    displayProducts(edtSearchProduct.Text);
end;

procedure TfrmMain.btnSeachProductClick(Sender: TObject);
begin
  displayProducts(edtSearchProduct.Text);
end;


procedure TfrmMain.SpeedButton2Click(Sender: TObject);
begin
  frmNewProduct :=  TfrmNewProduct.Create(Self);
    with frmNewProduct do
      begin
        try
          Hide;
          myProductMode := ProductUpdate;
          ShowModal;
        finally
          Free;
          displayProducts();
        end;
      end
end;

procedure TfrmMain.delete_product(ID: String);
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
      SQL.Add('set DEL_FLAG=''1''');
      SQL.Add('WHERE PRODUCT_ID=:ID');
      Parameters.ParamByName('ID').Value := ID;
      if not Prepared then ExecSQL;
    end;
  FreeAndNil(qry);
end;

procedure TfrmMain.delete_orders(ID: String);
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
      SQL.Add('UPDATE ORDERS');
      SQL.Add('set ORDER_DEL_FLAG=''1''');
      SQL.Add('WHERE TRIM(ORDER_NO)=:ID');
      Parameters.ParamByName('ID').Value := ID;
      if not Prepared then ExecSQL;
    end;
  FreeAndNil(qry);
end;

procedure TfrmMain.delete_releases(ID: String);
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
      SQL.Add('UPDATE release');
      SQL.Add('set RELEASE_DEL_FLAG=''1''');
      SQL.Add('WHERE TRIM(RELEASE_NO)=:ID');
      Parameters.ParamByName('ID').Value := TRIM(ID);
      if not Prepared then ExecSQL;
    end;
  FreeAndNil(qry);
end;

procedure TfrmMain.delete_sales(ID: String);
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
      SQL.Add('UPDATE sales');
      SQL.Add('set SALES_DEL_FLAG=''1''');
      SQL.Add('WHERE TRIM(SALES_NO)=:ID');
      Parameters.ParamByName('ID').Value := TRIM(ID);
      if not Prepared then ExecSQL;
    end;
  FreeAndNil(qry);
end;

procedure TfrmMain.SpeedButton3Click(Sender: TObject);
begin
  if DisplayMessage(mtConfirmation, [mbYes, mbNo], 3)  = mrYes then
    try
      DM.adoConn.BeginTrans;
        delete_product(dbProducts.Fields[0].AsString);
      DM.adoConn.CommitTrans;
      DisplayMessage(mtInformation, [mbOK], 4);
      ModalResult := mrOK;
    except
      DM.adoConn.RollbackTrans;
      DisplayMessage(mtError, [mbOK], 5);
    end;
  displayProducts();
end;
function TfrmMain.getProducts: TStrings;
var
  sl : TStringList;
  qry : TADOQuery;
begin
  qry := TADOQuery.Create(nil);
  with qry do
    begin
      Connection := DM.adoConn;
      //dbDisconnect;
      dbConnect(getConnString());
      SQL.Clear;
      SQL.Add('select PRODUCT_NAME & " " & SPECIFICATION from products');
      SQL.Add('where DEL_FLAG = ''0''');
      SQL.Add('order by PRODUCT_NAME & " " & SPECIFICATION');
      if not Prepared then Open;
      sl := TStringList.Create;
      if IsEmpty then
        sl.Add('')
      else
        begin
          while not EOF do
            begin
              sl.Add(Fields[0].AsString);
              Next;
            end;
        end;
    end;
  Result := sl;
  FreeAndNil(qry);
end;

function TfrmMain.check_products(NameSpecs : String) :  Boolean;
begin
  with DM.qrySelectProduct do
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
      SQL.Add('and TRIM(PRODUCT_NAME & " " & SPECIFICATION) =:NameSpecs');
      Parameters.ParamByName('NameSpecs').Value := Trim(NameSpecs);
      if not Prepared then Open;
      if IsEmpty then Result := False
      else Result := True;
    end;
end;

procedure TfrmMain.pageControlChange(Sender: TObject);
begin
  if pageControl.ActivePageIndex = 0 then        //products
    initProductpage
  else if pageControl.ActivePageIndex = 1 then   //orders
    initOrderspage
  else if pageControl.ActivePageIndex = 2 then   //release
    initReleasepage
  else if pageControl.ActivePageIndex = 3 then   //sales
    initSalespage
  else if pageControl.ActivePageIndex = 4 then   //report center
    btnReport.OnClick(Sender);
end;

procedure TfrmMain.initProductpage;
begin
  edtSearchProduct.Clear;
  displayProducts;
end;

procedure TfrmMain.initOrderspage;
begin
  dtpDelDate.DateTime := Now;
  cbOrderProduct.Clear;
  cbOrderProduct.Items := getProducts;
  edtOrderCode.Clear;
  edtOrderQty.Text := '0';
  edtOrderName.Clear;
  edtOrderSpecs.Clear;
  edtOrderStdCost.Clear;
  edtOrderSellPrice.Clear;
  lvOrders.Clear;
  dtpOrderFrom.DateTime := Now - 30;
  dtpOrderTo.DateTime := Now;
  edtOrderSearch.Clear;
  displayOrders(Trim(edtOrderSearch.Text));
  cbOrderProduct.SetFocus;
end;

procedure TfrmMain.cbOrderProductKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    begin
      if not check_products(Trim(cbOrderProduct.Text)) then
        SelectError(cbOrderProduct, 6)
      else
        begin
          with DM.qrySelectProduct do
            begin
              edtOrderCode.Text := Fields[0].AsString;
              edtOrderName.Text := Fields[1].AsString;
              edtOrderSpecs.Text := Fields[2].AsString;
              edtOrderStdCost.Text := Fields[3].AsString;
              edtOrderSellPrice.Text := Fields[4].AsString;
              edtOrderQty.SetFocus;
            end;
        end;
    end;
end;

procedure TfrmMain.edtOrderQtyKeyPress(Sender: TObject; var Key: Char);
var
  intBuff  : Integer;
begin
  if not (Key in ['0'..'9', #13, #8]) then Key :=#0
  else if Key =#13 then
    begin
      if Trim(edtOrderQty.Text) = '' then
        InputError(edtOrderQty, 2)
      else if not TryStrToInt(edtOrderQty.Text, intBuff) then
        InputError(edtOrderQty, 2)
      else if intbuff <= 0 then
        InputError(edtOrderQty, 2)
      else
        btnOrdersAdd.SetFocus
    end;

end;

procedure TfrmMain.btnProductSearchClick(Sender: TObject);
begin
    displayProducts(edtSearchProduct.Text);
end;

procedure TfrmMain.BitBtn2Click(Sender: TObject);
begin
  frmNewProduct :=  TfrmNewProduct.Create(Self);
    with frmNewProduct do
      begin
        try
          Hide;
          myProductMode := NewProduct;
          ShowModal;
        finally
          Free;
          displayProducts();
        end;
      end
end;

procedure TfrmMain.BitBtn4Click(Sender: TObject);
begin
  if DisplayMessage(mtConfirmation, [mbYes, mbNo], 3)  = mrYes then
    try
      DM.adoConn.BeginTrans;
        delete_product(dbProducts.Fields[0].AsString);
      DM.adoConn.CommitTrans;
      DisplayMessage(mtInformation, [mbOK], 4);
      ModalResult := mrOK;
    except
      DM.adoConn.RollbackTrans;
      DisplayMessage(mtError, [mbOK], 5);
    end;
  displayProducts();
end;

procedure TfrmMain.BitBtn3Click(Sender: TObject);
begin
  frmNewProduct :=  TfrmNewProduct.Create(Self);
    with frmNewProduct do
      begin
        try
          Hide;
          myProductMode := ProductUpdate;
          ShowModal;
        finally
          Free;
          displayProducts();
        end;
      end
end;

procedure TfrmMain.btnOrdersAddClick(Sender: TObject);
var
  c : Char;
  wk_delDate : String;
begin
  c := #13;
  FocusError := 0;
  if FocusError = 0 then cbOrderProduct.OnKeyPress(Sender, c);
  if FocusError = 0 then edtOrderQty.OnKeyPress(Sender, c);
  if FocusError = 0 then
    begin
      wk_delDate := FormatDateTime('YYYYMMDD', dtpDelDate.DateTime);
      lvOrders.Items.Add.Caption := wk_delDate;
      with DM.qrySelectProduct do
        begin
          lvOrders.Items.Item[lvOrders.Items.Count - 1].SubItems.Add(Fields[0].AsString);
          lvOrders.Items.Item[lvOrders.Items.Count - 1].SubItems.Add(Fields[1].AsString);
          lvOrders.Items.Item[lvOrders.Items.Count - 1].SubItems.Add(Fields[2].AsString);
          lvOrders.Items.Item[lvOrders.Items.Count - 1].SubItems.Add(edtOrderQty.Text);
          cbOrderProduct.SetFocus;
        end;
    end;
end;

procedure TfrmMain.pmOrderListPopup(Sender: TObject);
begin
  if (lvOrders.Items.Count = 0) then
    begin
      Remove.Enabled := not True;
      ClearAll.Enabled := not True;
    end
  else
    begin
      Remove.Enabled := True;
      ClearAll.Enabled := True;
    end;
end;

procedure TfrmMain.RemoveClick(Sender: TObject);
begin
  lvOrders.DeleteSelected;
  cbOrderProduct.SetFocus;
end;

procedure TfrmMain.ClearAllClick(Sender: TObject);
begin
  lvOrders.Clear;
  cbOrderProduct.SetFocus;
end;

procedure TfrmMain.btnOrdersPostClick(Sender: TObject);
var
  i : Integer;
  orderCode : String;
begin
  if lvOrders.Items.Count > 0 then
    begin
      if DisplayMessage(mtConfirmation, [mbYes, mbNo], 3)  = mrYes then
        try
          DM.adoConn.BeginTrans;
          orderCode := getNewOrderCode();
          for i:= 0 to lvOrders.Items.Count - 1 do
            begin
              insert_orders(orderCode,
                            lvOrders.Items[i].Caption,
                            lvOrders.Items.Item[i].SubItems[0],
                            lvOrders.Items.Item[i].SubItems[3],
                            '0');
            end;
          DM.adoConn.CommitTrans;
          DisplayMessage(mtInformation, [mbOK], 4);
          initOrderspage;
        except
          DM.adoConn.RollbackTrans;
          DisplayMessage(mtError, [mbOK], 5);
          initOrderspage
        end;
    end;
end;

procedure TfrmMain.displayOrders(aSearch: String);
var
  vSearch : String;
begin
  vSearch :=  '%' +  aSearch +  '%';
  with DM.qryOrders do
    begin
      Connection := DM.adoConn;
      //dbDisconnect;
      dbConnect(getConnString());
      SQL.Clear;
      SQL.Add('select ORDER_NO, ORDER_DATE, ORDER_PRODUCT_ID, PRODUCT_NAME, SPECIFICATION, ORDER_QUANTITY');
      SQL.Add('from orders , products');
      SQL.Add('where (ORDER_PRODUCT_ID = PRODUCT_ID)');
      SQL.Add('and ((PRODUCT_NAME like '+ QuotedStr(vSearch) + ') or (SPECIFICATION like ' + QuotedStr(vSearch) + ') or (ORDER_DATE like ' + QuotedStr(vSearch) +  '))' );
      SQL.Add('and ORDER_DEL_FLAG = ''0''');
      SQL.Add('order by ORDER_DATE desc');
      if not Prepared then Open;
    end;
end;

procedure TfrmMain.BitBtn1Click(Sender: TObject);
begin
  displayOrders(Trim(edtOrderSearch.Text));
  FitGrid(dbOrders);
  SetDBGridColumn(dbOrders);
end;

procedure TfrmMain.MenuItem1Click(Sender: TObject);
begin
  if trim(dbOrders.Fields[0].AsString) <> '' then
    begin
     if DisplayMessage(mtConfirmation, [mbYes, mbNo], 3)  = mrYes then
      try
        DM.adoConn.BeginTrans;
          delete_orders(dbOrders.Fields[0].AsString);
        DM.adoConn.CommitTrans;
        DisplayMessage(mtInformation, [mbOK], 4);
        ModalResult := mrOK;
      except
        DM.adoConn.RollbackTrans;
        DisplayMessage(mtError, [mbOK], 5);
      end;
    end;
  displayOrders(Trim(edtOrderSearch.Text))
end;

procedure TfrmMain.MenuItem2Click(Sender: TObject);
begin
  displayOrders(Trim(edtOrderSearch.Text));
end;

procedure TfrmMain.initReleasepage;
begin
  dtpRelDate.DateTime := Now;
  cbRelProduct.Clear;
  cbRelProduct.Items := getProducts;
  edtRelCode.Clear;
  edtRelQty.Text := '0';
  edtRelName.Clear;
  edtRelSpecs.Clear;
  edtRelInventory.Clear;
  lvRelease.Clear;
  dtpRelFrom.DateTime := Now - 30;
  dtpRelTo.DateTime := Now;
  edtRelSearch.Clear;
  displayReleases(Trim(edtRelSearch.Text));
  cbRelProduct.SetFocus;
end;

procedure TfrmMain.cbRelProductKeyPress(Sender: TObject; var Key: Char);
var
  wk_relInventory : Integer;
begin
  if Key = #13 then
    begin
      if not check_products(Trim(cbRelProduct.Text)) then
        SelectError(cbRelProduct, 6)
      else
        begin
          with DM.qrySelectProduct do
            begin
              wk_relInventory := getSumOrders(Fields[0].AsString) - getSumRelease(Fields[0].AsString);
              edtRelCode.Text := Fields[0].AsString;
              edtRelName.Text := Fields[1].AsString;
              edtRelSpecs.Text := Fields[2].AsString;
              edtRelInventory.Text:= IntToStr(wk_relInventory);
              edtRelQty.SetFocus;
            end;
        end;
    end;
end;

function TfrmMain.getSumOrders(ProductID : String) : Integer;
var
  qry : TADOQuery;
begin
  qry :=  TADOQuery.Create(nil);
  with qry do
    begin
      Connection := DM.adoConn;
      //dbDisconnect;
      dbConnect(getConnString());
      SQL.Add('select sum(ORDER_QUANTITY) from orders');
      SQL.Add('where ORDER_PRODUCT_ID =' + QuotedStr(ProductID));
      SQL.Add('and ORDER_DEL_FLAG  =''0''');
      if not Prepared then Open;
      if IsEmpty then
        Result := 0
      else
        Result := Fields[0].AsInteger;
    end;
  FreeAndNil(qry);
end;

function TfrmMain.getSumRelease(ProductID : String) : Integer;
var
  qry : TADOQuery;
begin
  qry :=  TADOQuery.Create(nil);
  with qry do
    begin
      Connection := DM.adoConn;
      //dbDisconnect;
      dbConnect(getConnString());
      SQL.Add('select sum(RELEASE_QUANTITY) from release');
      SQL.Add('where RELEASE_PRODUCT_ID =' + QuotedStr(ProductID));
      SQL.Add('and RELEASE_DEL_FLAG  =''0''');
      if not Prepared then Open;
      if IsEmpty then
        Result := 0
      else
        Result := Fields[0].AsInteger;
    end;
  FreeAndNil(qry);
end;

procedure TfrmMain.edtRelQtyKeyPress(Sender: TObject; var Key: Char);
var
  intBuff  : Integer;
begin
  if not (Key in ['0'..'9', #13, #8]) then Key :=#0
  else if Key =#13 then
    begin
      if Trim(edtRelQty.Text) = '' then
        InputError(edtRelQty, 2)
      else if not TryStrToInt(edtRelQty.Text, intBuff) then
        InputError(edtRelQty, 2)
      else if intbuff <= 0 then
        InputError(edtRelQty, 2)
      else if  intbuff > StrToInt(edtRelInventory.Text) then
        InputError(edtRelQty, 7)
      else
        btnRelAdd.SetFocus
    end;
end;

procedure TfrmMain.btnRelAddClick(Sender: TObject);
var
  c : Char;
  wk_RelDate : String;
begin
  c := #13;
  FocusError := 0;
  if FocusError = 0 then cbRelProduct.OnKeyPress(Sender, c);
  if FocusError = 0 then edtRelQty.OnKeyPress(Sender, c);
  if FocusError = 0 then
    begin
      wk_RelDate := FormatDateTime('YYYYMMDD', dtpRelDate.DateTime);
      lvRelease.Items.Add.Caption := wk_RelDate;
      with DM.qrySelectProduct do
        begin
          lvRelease.Items.Item[lvRelease.Items.Count - 1].SubItems.Add(Fields[0].AsString);
          lvRelease.Items.Item[lvRelease.Items.Count - 1].SubItems.Add(Fields[1].AsString);
          lvRelease.Items.Item[lvRelease.Items.Count - 1].SubItems.Add(Fields[2].AsString);
          lvRelease.Items.Item[lvRelease.Items.Count - 1].SubItems.Add(edtRelQty.Text);
          cbRelProduct.SetFocus;
        end;
    end;
end;

procedure TfrmMain.Remove1Click(Sender: TObject);
begin
  lvRelease.DeleteSelected;
  cbRelProduct.SetFocus;
end;

procedure TfrmMain.ClearAll1Click(Sender: TObject);
begin
  lvRelease.Clear;
  cbRelProduct.SetFocus;
end;

procedure TfrmMain.pmRelListPopup(Sender: TObject);
begin
  if (lvRelease.Items.Count = 0) then
    begin
      Remove1.Enabled := not True;
      ClearAll1.Enabled := not True;
    end
  else
    begin
      Remove1.Enabled := True;
      ClearAll1.Enabled := True;
    end;
end;

procedure TfrmMain.btnRelPostClick(Sender: TObject);
var
  i : Integer;
  RelCode : String;
begin
  if lvRelease.Items.Count > 0 then
    begin
     BranchCode := '';
     if ibx_branch.Execute <> 1 then
      BranchCode := ''
     else
      BranchCode := Copy(ibx_branch.InputText,1, 2);

    if BranchCode <> '' then
      begin
        if DisplayMessage(mtConfirmation, [mbYes, mbNo], 3)  = mrYes then
          try
            DM.adoConn.BeginTrans;
            RelCode := getNewRelCode();
            for i:= 0 to lvRelease.Items.Count - 1 do
              begin
                insert_release(RelCode,
                              lvRelease.Items[i].Caption,
                              lvRelease.Items.Item[i].SubItems[0],
                              lvRelease.Items.Item[i].SubItems[3],
                              BranchCode,
                              '0')
              end;
            DM.adoConn.CommitTrans;
            DisplayMessage(mtInformation, [mbOK], 4);
            initReleasepage;
          except
            DM.adoConn.RollbackTrans;
            DisplayMessage(mtError, [mbOK], 5);
            initReleasepage
          end;
      end;
    end;
end;

procedure TfrmMain.displayReleases(aSearch: String);
var
  vSearch : String;
begin
  vSearch :=  '%' +  aSearch +  '%';
  with DM.qryRelease do
    begin
      Connection := DM.adoConn;
      //dbDisconnect;
      dbConnect(getConnString());
      SQL.Clear;
      SQL.Add('select RELEASE_NO, RELEASE_DATE, RELEASE_PRODUCT_ID, PRODUCT_NAME, SPECIFICATION, format(LIST_PRICE,"Standard") as LIST_PRICE, RELEASE_QUANTITY, BRANCH_NAME');
      SQL.Add('from release , products, branches');
      SQL.Add('where (RELEASE_PRODUCT_ID = PRODUCT_ID)');
      SQL.Add('and RELEASE_BRANCH_CODE = BRANCH_CODE');
      SQL.Add('and ((PRODUCT_NAME like '+ QuotedStr(vSearch) + ') or (SPECIFICATION like ' + QuotedStr(vSearch) + ') or (RELEASE_DATE like '+  QuotedStr(vSearch)+'))' );
      SQL.Add('and TRIM(RELEASE_DEL_FLAG) = ''0''');
      SQL.Add('order by RELEASE_NO desc');
      if not Prepared then Open;
    end;
end;

procedure TfrmMain.displaySales(aSearch: String);
var
  vSearch : String;
begin
  vSearch :=  '%' +  aSearch +  '%';
  with DM.qrySales do
    begin
      Connection := DM.adoConn;
      //dbDisconnect;
      dbConnect(getConnString());
      SQL.Clear;
      SQL.Add('select SALES_NO, SALES_DATE, SALES_PRODUCT_ID, PRODUCT_NAME, SPECIFICATION, SALES_QUANTITY, format(SALES_SALE_PRICE, "Standard") as SALES_SALE_PRICE, BRANCH_NAME');
      SQL.Add('from sales , products, branches');
      SQL.Add('where (SALES_PRODUCT_ID = PRODUCT_ID)');
      SQL.Add('and SALES_BRANCH_CODE = BRANCH_CODE');
      SQL.Add('and ((PRODUCT_NAME like '+ QuotedStr(vSearch) + ') or (SPECIFICATION like ' + QuotedStr(vSearch) + ') or (SALES_DATE like ' + QuotedStr(vSearch) + '))' );
      SQL.Add('and TRIM(SALES_DEL_FLAG) = ''0''');
      SQL.Add('order by SALES_NO desc');
      if not Prepared then Open;
    end;
end;

procedure TfrmMain.BitBtn7Click(Sender: TObject);
begin
  displayReleases(Trim(edtRelSearch.Text));
  FitGrid(dbRelease);
  SetDBGridColumn(dbRelease);
end;

procedure TfrmMain.MenuItem3Click(Sender: TObject);
begin
  if trim(dbRelease.Fields[0].AsString) <> '' then
    begin
     if DisplayMessage(mtConfirmation, [mbYes, mbNo], 3)  = mrYes then
      try
        DM.adoConn.BeginTrans;
          delete_releases(dbRelease.Fields[0].AsString);
        DM.adoConn.CommitTrans;
        DisplayMessage(mtInformation, [mbOK], 4);
        ModalResult := mrOK;
      except
        DM.adoConn.RollbackTrans;
        DisplayMessage(mtError, [mbOK], 5);
      end;
    end;
  displayReleases(Trim(edtRelSearch.Text));
end;

procedure TfrmMain.MenuItem4Click(Sender: TObject);
begin
  displayOrders(Trim(edtOrderSearch.Text))
end;

procedure TfrmMain.initSalespage;
begin
  dtpSellDate.DateTime := Now;
  cbSellProduct.Clear;
  cbSellProduct.Items := getProducts;
  edtSellCode.Clear;
  edtSellQty.Text := '0';
  edtSellPrice.Text := '0';
  edtBranchInventory.Text := '0';
  edtSellName.Clear;
  edtSellSpecs.Clear;
  lvSell.Clear;
  edtOrderSearch.Clear;
  displaySales(trim(edtSellSearch.Text));
  cbBranch.SetFocus;
  InputChecker := False;
end;

procedure TfrmMain.cbBranchChange(Sender: TObject);
begin
  cbSellProduct.SetFocus;
end;

procedure TfrmMain.cbSellProductKeyPress(Sender: TObject; var Key: Char);
var
   wk_sellInventory : Integer;
begin
  if Key = #13 then
    begin
      if not check_products(Trim(cbSellProduct.Text)) then
        SelectError(cbSellProduct, 6)
      else
        begin
          with DM.qrySelectProduct do
            begin

              wk_sellInventory := getSumBranchRelease(Fields[0].AsString, Copy(cbBranch.Text, 1, 2))
                                      - getSumBranchSalesProduct(Fields[0].AsString, Copy(cbBranch.Text, 1, 2)) ;
              edtSellCode.Text := Fields[0].AsString;
              edtSellName.Text := Fields[1].AsString;
              edtSellSpecs.Text := Fields[2].AsString;
              if not InputChecker then
                edtSellPrice.Text := Fields[4].AsString;
              edtBranchInventory.Text:= IntToStr(wk_sellInventory);
              edtSellQty.SetFocus;
            end;
        end;
    end;
end;


function TfrmMain.getSumBranchRelease(ProductID,
  BranchID: String): Integer;
var
  qry : TADOQuery;
begin
  qry :=  TADOQuery.Create(nil);
  with qry do
    begin
      Connection := DM.adoConn;
      //dbDisconnect;
      dbConnect(getConnString());
      SQL.Add('select sum(RELEASE_QUANTITY) from release');
      SQL.Add('where ((RELEASE_PRODUCT_ID =' + QuotedStr(ProductID));
      SQL.Add('and RELEASE_BRANCH_CODE=' + QuotedStr(BranchID));
      SQL.Add(')and RELEASE_DEL_FLAG  =''0'')');
      if not Prepared then Open;
      if IsEmpty then
        Result := 0
      else
        Result := Fields[0].AsInteger;
    end;
  FreeAndNil(qry);
end;

function TfrmMain.getSumBranchSalesProduct(ProductID, BranchID: String): Integer;
var
  qry : TADOQuery;
begin
  qry :=  TADOQuery.Create(nil);
  with qry do
    begin
      Connection := DM.adoConn;
      //dbDisconnect;
      dbConnect(getConnString());
      SQL.Add('select sum(SALES_QUANTITY) from sales');
      SQL.Add('where ((SALES_PRODUCT_ID =' + QuotedStr(ProductID));
      SQL.Add('and SALES_BRANCH_CODE=' + QuotedStr(BranchID));
      SQL.Add(')and SALES_DEL_FLAG  =''0'')');
      if not Prepared then Open;
      if IsEmpty then
        Result := 0
      else
        Result := Fields[0].AsInteger;
    end;
  FreeAndNil(qry);
end;

procedure TfrmMain.edtSellQtyKeyPress(Sender: TObject; var Key: Char);
var
  intBuff  : Integer;
begin
  if not (Key in ['0'..'9', #13, #8]) then Key :=#0
  else if Key =#13 then
    begin
      if Trim(edtSellQty.Text) = '' then
        InputError(edtSellQty, 2)
      else if not TryStrToInt(edtSellQty.Text, intBuff) then
        InputError(edtSellQty, 2)
      else if intbuff <= 0 then
        InputError(edtSellQty, 2)
      else if  intbuff > StrToInt(edtBranchInventory.Text) then
        InputError(edtSellQty, 7)
      else
        edtSellPrice.SetFocus
    end;
end;

procedure TfrmMain.edtSellPriceKeyPress(Sender: TObject; var Key: Char);
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
           btnSellAdd.Setfocus;
        end;
    end;
end;
procedure TfrmMain.btnSellAddClick(Sender: TObject);
var
  c : Char;
  wk_SellDate : String;
begin
  c := #13;
  FocusError := 0;
  InputChecker := True;
  if FocusError = 0 then cbSellProduct.OnKeyPress(Sender, c);
  if FocusError = 0 then edtSellQty.OnKeyPress(Sender, c);
  if FocusError = 0 then edtSellPrice.OnKeyPress(Sender, c);
  if FocusError = 0 then
    begin
      wk_SellDate := FormatDateTime('YYYYMMDD', dtpSellDate.DateTime);
      lvSell.Items.Add.Caption := wk_SellDate;
      with DM.qrySelectProduct do
        begin
          lvSell.Items.Item[lvSell.Items.Count - 1].SubItems.Add(Fields[0].AsString);
          lvSell.Items.Item[lvSell.Items.Count - 1].SubItems.Add(Fields[1].AsString);
          lvSell.Items.Item[lvSell.Items.Count - 1].SubItems.Add(Fields[2].AsString);
          lvSell.Items.Item[lvSell.Items.Count - 1].SubItems.Add(edtSellQty.Text);
          lvSell.Items.Item[lvSell.Items.Count - 1].SubItems.Add(edtSellPrice.Text);
          lvSell.Items.Item[lvSell.Items.Count - 1].SubItems.Add(cbBranch.Text);
          cbSellProduct.SetFocus;
        end;
    end;
  InputChecker := False;
end;

procedure TfrmMain.BitBtn6Click(Sender: TObject);
var
  i : Integer;
  SellCode : String;
begin
  if lvSell.Items.Count > 0 then
    begin
      if DisplayMessage(mtConfirmation, [mbYes, mbNo], 3)  = mrYes then
        try
          DM.adoConn.BeginTrans;
          SellCode := getNewSalesCode();
          for i:= 0 to lvSell.Items.Count - 1 do
            begin
              insert_sales(SellCode,
                            lvSell.Items[i].Caption,
                            lvSell.Items.Item[i].SubItems[0],
                            lvSell.Items.Item[i].SubItems[3],
                            StringReplace(lvSell.Items.Item[i].SubItems[4], ',', '', [rfReplaceAll, rfIgnoreCase]),
                            Copy(lvSell.Items.Item[i].SubItems[5], 1, 2),
                            '0')
            end;
          DM.adoConn.CommitTrans;
          DisplayMessage(mtInformation, [mbOK], 4);
          initSalespage;
        except
          DM.adoConn.RollbackTrans;
          DisplayMessage(mtError, [mbOK], 5);
          initSalespage
        end;    
    end;
end;

procedure TfrmMain.btnSellSearchClick(Sender: TObject);
begin
  displaySales(trim(edtSellSearch.Text));
  FitGrid(dbSell);
  SetDBGridColumn(dbSell);
end;

procedure TfrmMain.ClearAll2Click(Sender: TObject);
begin
  lvSell.Clear;
  cbSellProduct.SetFocus;
end;

procedure TfrmMain.Remove2Click(Sender: TObject);
begin
  lvSell.DeleteSelected;
  cbSellProduct.SetFocus;
end;

procedure TfrmMain.pmSellListPopup(Sender: TObject);
begin
  if (lvSell.Items.Count = 0) then
    begin
      Remove2.Enabled := not True;
      ClearAll2.Enabled := not True;
    end
  else
    begin
      Remove2.Enabled := True;
      ClearAll2.Enabled := True;
    end;
end;

procedure TfrmMain.MenuItem5Click(Sender: TObject);
begin
  if trim(dbsELL.Fields[0].AsString) <> '' then
    begin
     if DisplayMessage(mtConfirmation, [mbYes, mbNo], 3)  = mrYes then
      try
        DM.adoConn.BeginTrans;
          delete_sales(dbSell.Fields[0].AsString);
        DM.adoConn.CommitTrans;
        DisplayMessage(mtInformation, [mbOK], 4);
        ModalResult := mrOK;
      except
        DM.adoConn.RollbackTrans;
        DisplayMessage(mtError, [mbOK], 5);
      end;
    end;
  displaySales(Trim(edtSellSearch.Text));
end;

procedure TfrmMain.MenuItem6Click(Sender: TObject);
begin
  displaySales(trim(edtSellSearch.Text));
end;

procedure TfrmMain.edtOrderSearchKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    displayOrders(Trim(edtOrderSearch.Text));
end;

procedure TfrmMain.edtRelSearchKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    displayReleases(Trim(edtRelSearch.Text));
end;

procedure TfrmMain.edtSellSearchKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    displaySales(trim(edtSellSearch.Text));
end;

procedure TfrmMain.getWarehouseInventory(ASearch: String);
var
  mySQL, vSearch : String;
  i : Integer;
begin;
  aBsy_main.Show;
  getCDSProducts(ASearch);
  if DM.qryCDSProducts.RecordCount > 0 then
    begin
      cdsWarehouseInv :=  TClientDataSet.Create(Application);
      with cdsWarehouseInv do     //Add columns
        begin
          Fields.Clear;
          FieldDefs.Add('ProductID', ftString, 50, False);
          FieldDefs.Add('ProductName', ftString, 50, False);
          FieldDefs.Add('Specification', ftString, 50, False);
          FieldDefs.Add('TotalOrders', ftString, 50, False);
          FieldDefs.Add('TotalReleases', ftString, 50, False);
          FieldDefs.Add('StockQuantity', ftString, 50, False);
          CreateDataSet;
          Open;
        end;
      with DM.qryCDSProducts do  //Add Rows
        begin
          First;
          i := 0;
          aBsy_main.ProgressMax := DM.qryCDSProducts.RecordCount;
          while not EOF do
            begin
              with cdsWarehouseInv do
                begin
                  Append;
                  FieldByName('ProductID').AsString := DM.qryCDSProducts.FieldByName('PRODUCT_ID').AsString;
                  FieldByName('ProductName').AsString := DM.qryCDSProducts.FieldByName('PRODUCT_NAME').AsString;
                  FieldByName('Specification').AsString := DM.qryCDSProducts.FieldByName('PRODUCT_NAME').AsString;
                  FieldByName('TotalOrders').AsString := IntToStr(getSumOrders(DM.qryCDSProducts.FieldByName('PRODUCT_ID').AsString));
                  FieldByName('TotalReleases').AsString := IntToStr(getSumRelease(DM.qryCDSProducts.FieldByName('PRODUCT_ID').AsString));
                  FieldByName('StockQuantity').AsString := IntToStr(getSumOrders(DM.qryCDSProducts.FieldByName('PRODUCT_ID').AsString) -
                                                                getSumRelease(DM.qryCDSProducts.FieldByName('PRODUCT_ID').AsString));

                  Post;
                  Inc(i);
                  aBsy_main.ProgressPos := i;
                end;
              Next;
            end;
        end;
      DM.dtsCDSProducts.DataSet := cdsWarehouseInv;
      dbReport.DataSource := DM.dtsCDSProducts;
    end;
   aBsy_main.Hide;
end;

procedure TfrmMain.btnReportClick(Sender: TObject);
begin
  case cbReportType.ItemIndex of
    0 : getWarehouseInventory(Trim(edtReportSearch.Text));
    1 : getBranchInventory('01', Trim(edtReportSearch.Text));
    2 : getBranchInventory('02', Trim(edtReportSearch.Text));
    3 : getSales(Trim(edtReportSearch.Text));
  end;
  FitGrid(dbReport);
  SetDBGridColumn(dbReport);
end;

procedure TfrmMain.FitGrid(Grid: TDBGrid);
const
  C_Add=3;
var
  ds: TDataSet;
  bm: TBookmark;
  i: Integer;
  w: Integer;
  a: Array of Integer;
begin
  ds := Grid.DataSource.DataSet;
  if Assigned(ds) then
  begin
    ds.DisableControls;
    bm := ds.GetBookmark;
    try
      ds.First;
      SetLength(a, Grid.Columns.Count);
      while not ds.Eof do
      begin
        for I := 0 to Grid.Columns.Count - 1 do
        begin
          if Assigned(Grid.Columns[i].Field) then
          begin
            w :=  Grid.Canvas.TextWidth(ds.FieldByName(Grid.Columns[i].Field.FieldName).DisplayText);
            if a[i] < w  then
               a[i] := w ;
          end;
        end;
        ds.Next;
      end;
      for I := 0 to Grid.Columns.Count - 1 do
        Grid.Columns[i].Width := a[i] + C_Add;
        ds.GotoBookmark(bm);
    finally
      ds.FreeBookmark(bm);
      ds.EnableControls;
    end;
  end;
end;




procedure TfrmMain.SetDBGridColumn(Grid: TDBGrid);
var
  i:integer;
  wid:integer;
begin
  for i:=0 to Grid.Columns.Count-1 do
    begin
      wid:=Grid.Canvas.TextWidth(Grid.Columns.Items[i].Title.Caption);
      if wid > Grid.Columns.Items[i].Width then
        Grid.Columns.Items[i].Width:=wid+10;
    end;
end;
procedure TfrmMain.getBranchInventory(branchCode, ASearch: String);
var
  mySQL, vSearch : String;

begin
  vSearch :=  '%' +  ASearch +  '%';
  mySQL := 'select  ProductID, ProductName, Specification, TotalRelease, TotalSales, iif (IsNull(TotalRelease - TotalSales), ''0'', TotalRelease - TotalSales) as StockQuantity from ' +
           '(SELECT products.PRODUCT_ID as ProductID, products.PRODUCT_NAME as ProductName, '  +
           ' products.SPECIFICATION As Specification, release.RELEASE_BRANCH_CODE as RelBrCode, sales.SALES_BRANCH_CODE as SalesBrCode,' +
           ' iif (IsNull(Sum(release.RELEASE_QUANTITY)), ''0'',Sum(release.RELEASE_QUANTITY)) as TotalRelease, '  +
           ' iif (IsNull(Sum(sales.SALES_QUANTITY)), ''0'',Sum(sales.SALES_QUANTITY)) as TotalSales '  +
           //' iif (IsNull(Sum(release.RELEASE_QUANTITY) - Sum(sales.SALES_QUANTITY)), ''0'', Sum(release.RELEASE_QUANTITY) - Sum(sales.SALES_QUANTITY)) as StockQuantity  '  +
           ' FROM ((products LEFT OUTER JOIN release ON (products.PRODUCT_ID = release.RELEASE_PRODUCT_ID AND release.RELEASE_DEL_FLAG = ''0'' AND release.RELEASE_BRANCH_CODE ='+ QuotedStr(branchCode) + ')) ' +
           '       LEFT OUTER JOIN sales ON (products.PRODUCT_ID = sales.SALES_PRODUCT_ID AND SALES.SALES_DEL_FLAG = ''0'' AND sales.SALES_BRANCH_CODE =' + QuotedStr(branchCode) +'))';


  with DM.qryReport do
    begin
      Connection := DM.adoConn;
      //dbDisconnect;
      dbConnect(getConnString());
      SQL.Clear;
      SQL.Add(mySQL);
      SQL.Add('GROUP BY products.PRODUCT_ID, products.PRODUCT_NAME, products.SPECIFICATION, release.RELEASE_BRANCH_CODE, sales.SALES_BRANCH_CODE)');
      SQL.Add('where ((ProductName like '+ QuotedStr(vSearch) + ') or (Specification like ' + QuotedStr(vSearch) + '))' );
      if not Prepared then Open;
    end;
  dbReport.DataSource := DM.dtsReport;
end;

procedure TfrmMain.BitBtn8Click(Sender: TObject);
var
  excelDataSetQuery : TDataset;
begin
  case cbReportType.ItemIndex of
    0 :  excelDataSetQuery := cdsWarehouseInv
  else
    excelDataSetQuery := DM.qryReport;
  end;

  if excelDataSetQuery.Active then
  begin
    with scExcelExport do
      try
        aBsy_main.Show;
        Dataset := excelDataSetQuery;
        WorksheetName := cbReportType.Text;
        ExportDataset;
        aBsy_main.Hide;
      finally
        Disconnect;
      end;
  end;

end;

procedure TfrmMain.getSales(ASearch: String);
var
  mySQL, vSearch : String;

begin
  vSearch :=  '%' +  ASearch +  '%';
  mySQL :=
          'SELECT sales.SALES_DATE as SalesDate, products.PRODUCT_ID as ProductID, products.PRODUCT_NAME as ProductName, products.SPECIFICATION as Specification, sales.SALES_QUANTITY as SalesQuantity, ' +
          ' format(sales.SALES_SALE_PRICE, "Standard") as SalePrice, format((sales.SALES_QUANTITY *  sales.SALES_SALE_PRICE), "Standard") as Amount ,branches.BRANCH_NAME as Branch ' +
          'FROM (products INNER JOIN sales ON (products.PRODUCT_ID = sales.SALES_PRODUCT_ID AND sales.SALES_DEL_FLAG = ''0'')) LEFT JOIN branches ON sales.SALES_BRANCH_CODE = branches.BRANCH_CODE';


  with DM.qryReport do
    begin
      Connection := DM.adoConn;
      //dbDisconnect;
      dbConnect(getConnString());
      SQL.Clear;
      SQL.Add(mySQL);
      SQL.Add('where ((products.PRODUCT_NAME like '+ QuotedStr(vSearch) + ') or (products.SPECIFICATION like ' + QuotedStr(vSearch) + ') or (sales.SALES_DATE like ' + QuotedStr(vSearch) + '))' );
      if not Prepared then Open;
    end;
  dbReport.DataSource := DM.dtsReport;
end;

procedure TfrmMain.ExportOrdersClick(Sender: TObject);
begin
  if DM.qryOrders.Active then
    begin
      with scExcelExport do
        try
          aBsy_main.Show;
          Dataset := DM.qryOrders;
          WorksheetName := 'Orders';
          ExportDataset;
          aBsy_main.Hide;
        finally
          Disconnect;
        end;
    end;
end;

procedure TfrmMain.ExportToExcel1Click(Sender: TObject);
begin
  if DM.qryRelease.Active then
    begin
      with scExcelExport do
        try
          aBsy_main.Show;
          Dataset := DM.qryRelease;
          WorksheetName := 'Releases';
          ExportDataset;
          aBsy_main.Hide;
        finally
          Disconnect;
        end;
    end;
end;

procedure TfrmMain.ExportToExcel2Click(Sender: TObject);
begin
  if DM.qrySales.Active then
    begin
      with scExcelExport do
        try
          aBsy_main.Show;
          Dataset := DM.qrySales;
          WorksheetName := 'Sales';
          ExportDataset;
          aBsy_main.Hide;
        finally
          Disconnect;
        end;
    end;
end;

procedure TfrmMain.ExportToExcel3Click(Sender: TObject);
begin
  if DM.qryProducts.Active then
    begin
      with scExcelExport do
        try
          aBsy_main.Show;
          Dataset := DM.qryProducts;
          WorksheetName := 'Products';
          ExportDataset;
          aBsy_main.Hide;
        finally
          Disconnect;
        end;
    end;
end;

end.
