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

'販売数だけ商品棚から消す
'@return 販売数
Public Function SellItems() As Integer

    Dim i As Integer
    Dim item_ As Item
    For i = m_salesNum To 1 Step -1
        Set item_ = m_items.Item(i)
        m_items.Remove (i)
        SellItems = SellItems + 1
    Next i
    
    PrintItems
End Function
'当日の販売金額をカウント
'@return 販売金額
Public Function CountSalesPrice() As Integer

    Dim i As Integer
    Dim item_ As Item
    For i = DecideSales To 1 Step -1
        Set item_ = m_items.Item(i)
        CountSalesPrice = CountSalesPrice + item_.GetPrice '販売金額を加算
    Next i
      
End Function
'当日の販売数を算出する。とりあえずランダムな値を出しているだけ。
Private Function DecideSales() As Integer
    Dim max As Integer
    Dim min As Integer
    Dim ws As Worksheet
    Set ws = Worksheets("data")

    max = ws.Cells(4, 10)
    min = ws.Cells(3, 10)
    
    Randomize
    m_salesNum = Int((max - min + 1) * Rnd + min)
    DecideSales = m_salesNum
    
End Function

'当日が賞味期限の商品を２０円値引き
'@return 値引き数
Public Function PriceCutItems() As Integer
    
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
  PriceCutItems = priceCutNum
  PrintItems
  Debug.Print "Cabinet:PriceCutItems$ priceCut=", priceCutNum
End Function

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
  Debug.Print "Cabinet:DiscardItems$ loss=", discardNum
End Function


'現在の在庫状況をMainシートに転記する
Private Sub PrintItems()
    Dim ws As Worksheet
    Set ws = Worksheets("main")

    Dim i As Integer
    
    For i = 1 To 100
        ws.Cells(MAIN_ROW.ITEMS + i, MAIN_COLUMN.Number).ClearContents
    Next i
    
    Dim item_ As Item
    For i = 1 To m_items.Count
        Set item_ = m_items.Item(i)
        ws.Cells(MAIN_ROW.ITEMS + i, MAIN_COLUMN.Number) = item_.GetLimit
        ws.Cells(MAIN_ROW.ITEMS + i, MAIN_COLUMN.Price) = item_.GetPrice
    Next i
End Sub


