VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'商品を表すクラス。賞味期限と標準売価はコンストラクタでのみ設定可能。値引きのみセッターをつける。

Private m_limit As Date '賞味期限
Private m_price As Integer

Public Sub Class_Initialize()
    m_price = 100
    m_limit = Main.GetIndicator.GetRecord.GetDate + 2
End Sub


Public Function GetLimit()
    GetLimit = m_limit
End Function

Public Function GetPrice()
    GetPrice = m_price
End Function
Public Sub PriceDown(down_)
    m_price = m_price - down_
End Sub
