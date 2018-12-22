VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cabinet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'商品棚を表すクラス。Itemnの集合を表すコレクションを内包し、納品、販売、値引き、廃棄等を行うメソッドをもつ。

Private m_items As Collection
Private m_salesNum As Integer
Private m_salesPrice As Integer
Private m_priceCutNum As Integer
Private id As Integer


Private Sub Class_Initialize()
    Set m_items = New Collection
    id = 0
End Sub


'発注された商品を生成して棚に並べる
'本来であれば、生成は別のクラス/メソッドでやるべきな気がするがクラス数を抑える為に取り敢えずここで纏めている。
Public Sub DisplayItems(num As Integer)

    Dim i As Integer
    For i = 0 To num - 1
        Dim item_ As Item
        Set item_ = New Item
        m_items.Add item_, CStr(id)
        id = id + 1
    Next i
        PrintItems
End Sub


'当日の販売数を算出する。
Public Function DecideSales() As Integer
 
    Dim Customers As Collection
    Dim customerNum As Integer
    
    Const max = 10
    Const min = 5
    Randomize
    customerNum = Int((max - min + 1) * Rnd + min)
    m_salesPrice = 0 '毎回初期化しないと前の日の分が残ったままになる。
    Dim decideSales_ As Integer
    Dim i As Integer
    For i = 0 To customerNum
        Dim customer_ As Customer
        Set customer_ = New Customer
        decideSales_ = decideSales_ + customer_.BuyItems
        m_salesPrice = m_salesPrice + customer_.GetPay '一人当たりの支払いを人数分、合計してメンバ変数に保存。
    Next i
    DecideSales = decideSales_
    
End Function

'当日が賞味期限の商品を２０円値引きして、結果をメンバ変数に保存する。
'他クラスから値引き数を呼び出すときはfunc GetPriceCutNum()を呼ぶ。(これをそのまま２回呼ぶとバグるため）
Public Sub PriceCutItems()

  Dim priceCutNum As Integer
  Dim i As Integer
  Dim item_ As Item
  For i = m_items.Count - 1 To 1 Step -1
    Set item_ = m_items.Item(i)
    If item_.GetLimit = Main.GetIndicator.GetRecord.GetDate Then
        item_.PriceDown (20)
        priceCutNum = priceCutNum + 1
    End If
  Next i
  m_priceCutNum = priceCutNum
  PrintItems
End Sub

'当日が賞味期限の商品を廃棄
'@return 廃棄数
Public Function DiscardItems() As Integer
    
  Dim discardNum As Integer
  Dim i As Integer
  Dim item_ As Item
  For i = m_items.Count - 1 To 1 Step -1
    Set item_ = m_items.Item(i)
    If item_.GetLimit = Main.GetIndicator.GetRecord.GetDate Then
        m_items.Remove (i)
        discardNum = discardNum + 1
    End If
  Next i
  DiscardItems = discardNum
  PrintItems
End Function


'現在の在庫状況をMainシートに転記する
Private Sub PrintItems()
    Dim ws As Worksheet
    Set ws = Worksheets("main")

    Dim i As Integer
    
    For i = 1 To 100
        ws.Cells(MAIN_ROW.ITEMS + i, MAIN_COLUMN.number).ClearContents
    Next i
    
    Dim item_ As Item
    For i = 1 To m_items.Count
        Set item_ = m_items.Item(i)
        ws.Cells(MAIN_ROW.ITEMS + i, MAIN_COLUMN.number) = item_.GetLimit
        ws.Cells(MAIN_ROW.ITEMS + i, MAIN_COLUMN.Price) = item_.GetPrice
    Next i
End Sub

Public Function GetItems() As Collection
    Set GetItems = m_items
End Function
Public Function GetSalesPrice() As Integer
    GetSalesPrice = m_salesPrice
End Function
Public Function GetPriceCutNum() As Integer
    GetPriceCutNum = m_priceCutNum
End Function

