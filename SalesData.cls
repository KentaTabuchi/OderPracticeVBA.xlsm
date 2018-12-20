VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SalesData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'販売実績データのレコード一列分を表すクラス。

Private m_date As Date
Private m_delivery As Integer
Private m_sales As Integer
Private m_salesPrice As Integer
Private m_loss As Integer
Private m_stock As Integer
Private m_oder As Integer
Private m_priceCut As Integer

Public Function GetDate()
    GetDate = m_date
End Function
Public Sub SetDate(DATE_ As Date)
    m_date = DATE_
End Sub

Public Function GetDelivery()
    GetDelivery = m_delivery
End Function
Public Sub SetDelivery(delivery_ As Integer)
    m_delivery = delivery_
End Sub

Public Function GetSales() As Integer
    GetSales = m_sales
End Function
Public Sub SetSales(sales_ As Integer)
    m_sales = sales_
End Sub
Public Function GetSalesPrice() As Integer
    GetSalesPrice = m_salesPrice
End Function
Public Sub SetSalesPrice(salesPrice_ As Integer)
    m_salesPrice = salesPrice_
End Sub

Public Function GetLoss() As Integer
    GetLoss = m_loss
End Function

Public Sub SetLoss(loss_ As Integer)
    m_loss = loss_
End Sub
Public Function GetStock() As Integer
    GetStock = m_stock
End Function

Public Sub SetStock(stock_ As Integer)
    m_stock = stock_
End Sub
Public Function GetOder() As Integer
    GetOder = m_oder
End Function

Public Sub SetOder(oder_ As Integer)
    m_oder = oder_
End Sub

Public Function GetPriceCut() As Integer
    GetPriceCut = m_priceCut
End Function

Public Sub SetPriceCut(priceCut_ As Integer)
    m_priceCut = priceCut_
End Sub

