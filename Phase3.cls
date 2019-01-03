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
    
    Dim record As SalesData
    Set record = Main.GetIndicator.GetRecord
    
    Dim ws As Worksheet
    Set ws = Worksheets("main")
    
    record.SetLoss (Main.GetCabinet.DiscardItems)
    MessagePrinter.WriteLine record.GetLoss & "個廃棄しました。"
    Main.GetIndicator.UpdateRecord loss_:=record.GetLoss
    
    record.SetStock (record.GetStock - record.GetLoss)
    If (record.GetStock < 0) Then
        record.SetStock (0)
    End If
    
    Main.GetIndicator.UpdateRecord stock_:=record.GetStock
    
    
    TablePrinter.PrintCurrentRecordToTable record
    
    'TODO:ここに値引き後の在庫金額を書く。正値在庫＋値引き在庫
    
     MessagePrinter.WriteLine Strings.PLEASE_THE_ODER, 4
End Sub

Public Sub IPhase_ChangePhase()
    Main.SetCurrentPhase (PhaseNumber.LAST_4)
    MainSheet.Cells(MAIN_PHASE_PAIN.ROW, MAIN_PHASE_PAIN.COLUMN) = Strings.PHASE_ODER
End Sub

