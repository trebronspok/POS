object DM: TDM
  OldCreateOrder = False
  Left = 533
  Top = 157
  Height = 727
  Width = 1280
  object adoConn: TADOConnection
    LoginPrompt = False
    Mode = cmShareDenyNone
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 40
    Top = 24
  end
  object dtsProduct: TDataSource
    DataSet = qryProducts
    Left = 108
    Top = 92
  end
  object qryProducts: TADOQuery
    Connection = adoConn
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from products')
    Left = 44
    Top = 92
  end
  object qrySelectProduct: TADOQuery
    Connection = adoConn
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from products')
    Left = 48
    Top = 144
  end
  object qryOrders: TADOQuery
    Connection = adoConn
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from products')
    Left = 44
    Top = 208
  end
  object dtsOrders: TDataSource
    DataSet = qryOrders
    Left = 104
    Top = 212
  end
  object qryRelease: TADOQuery
    Connection = adoConn
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from products')
    Left = 44
    Top = 268
  end
  object dtsRelease: TDataSource
    DataSet = qryRelease
    Left = 108
    Top = 268
  end
  object qrySales: TADOQuery
    Connection = adoConn
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from products')
    Left = 48
    Top = 332
  end
  object dtsSales: TDataSource
    DataSet = qrySales
    Left = 108
    Top = 332
  end
  object qryReport: TADOQuery
    Connection = adoConn
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from products')
    Left = 52
    Top = 392
  end
  object dtsReport: TDataSource
    DataSet = qryReport
    Left = 120
    Top = 392
  end
  object qryCDSProducts: TADOQuery
    Connection = adoConn
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from products')
    Left = 164
    Top = 456
  end
  object dtsCDSProducts: TDataSource
    DataSet = cdsWarehouseInv
    Left = 260
    Top = 464
  end
  object cdsWarehouseInv: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 96
    Top = 536
  end
  object dspWarehouseInv: TDataSetProvider
    DataSet = cdsWarehouseInv
    Constraints = True
    Left = 236
    Top = 532
  end
end
