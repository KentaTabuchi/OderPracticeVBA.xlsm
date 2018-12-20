VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Phase3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'廃棄フェイズ

Implements IPhase

Public Sub IPhase_ExecutePhase()
    Debug.Print "phase:LOSS_3"
    
    Dim record As SalesData
    Set record = Main.GetIndicator.GetRecord
    
    Dim ws As Worksheet
    Set ws = Worksheets("main")
    
    record.SetLoss (Main.GetCabinet.DiscardItems)
    Main.GetIndicator.UpdateRecord loss_:=record.GetLoss
    
    record.SetStock (record.GetStock - record.GetLoss)
    If (record.GetStock < 0) Then
        record.SetStock (0)
    End If
    
    ws.Cells(MAIN_ROW.loss, MAIN_COLUMN.Number) = record.GetLoss
    ws.Cells(MAIN_ROW.loss, MAIN_COLUMN.Price) = record.GetLoss * -100
    
    'ws.Cells(MAIN_ROW.CURRENT_STOCK, MAIN_COLUMN.Number) = record.GetStock
    
    'ここに値引き後の在庫金額を書く。正値在庫＋値引き在庫
    
     Message.WriteLine Strings.str_PleaseTheOder, 4
End Sub

Public Sub IPhase_ChangePhase()
    Main.SetCurrentPhase (PhaseNumber.LAST_4)
End Sub

