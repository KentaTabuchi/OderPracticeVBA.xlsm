VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Customers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'現状仮に違うクラスでランダム値を使っているがこちらでAI化する。
'お客様を表すクラス買い物をする。販売数の決定を行う。
'お客様は曜日、天候によって人数がかわる。
'その日棚にある商品以上は買えない。
'購入意欲が規定ポイント以上だと買う。
'一つ買うごとに購入意欲は下がる。
'値引きしていると購入意欲が上がる。

Private number As Integer '客数
Private m_motivation As Integer '購入意欲

Private Sub Class_Initialize()
    m_motivation = 30
    'TODO:ランダム処理
End Sub
'@return 購入数
Public Function BuyItems() As Integer
    
    Dim buyItems_ As Integer
    
    Dim i As Integer
    Dim item_ As Item
    Dim cabinet_ As Cabinet: Set cabinet_ = Main.GetCabinet
    
    For i = cabinet_.GetItems.Count - 1 To 1 Step -1
        Set item_ = m_items.Item(i)
        
        If IsBuy(item_) Then
            buyItems_ = buyItems_ + 1
            m_motivation = m_motivation - 10
        End If
  
    Next i

    
End Function

'モチベーションが一定量を超えると買う
'@return 買うかどうか
Private Function IsBuy(item_ As Item) As Boolean

    If item_.GetPrice < 100 Then
        m_motivation = m_motivation + 20
    End If
    
    If m_motivation > 50 Then
        IsBuy = True
    End If
    
End Function


