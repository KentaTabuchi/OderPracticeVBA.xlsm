VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Phase1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'納品フェイズ

Implements IPhase

Public Sub IPhase_ExecutePhase()
    Debug.Print "phase:DELIVERY_1"
    
    Dim record As SalesData
    Set record = Main.GetIndicator.GetRecord
    
    Dim ws As Worksheet
    Set ws = Worksheets("main")
    
    record.SetStock (record.GetStock + record.GetDelivery)
    
    Main.GetCabinet.DisplayItems (Main.GetIndicator.GetRecord.GetDelivery)
    
    ws.Cells(MAIN_ROW.DELIVERY, MAIN_COLUMN.Number) = record.GetDelivery
    ws.Cells(MAIN_ROW.DELIVERY, MAIN_COLUMN.Price) = record.GetDelivery * 100
    
    'ws.Cells(MAIN_ROW.CURRENT_STOCK, MAIN_COLUMN.Number) = record.GetStock
    'ws.Cells(MAIN_ROW.CURRENT_STOCK, MAIN_COLUMN.Price) = record.GetStock * 100

    
End Sub

Public Sub IPhase_ChangePhase()
    Main.SetCurrentPhase (PhaseNumber.SELL_2)
End Sub
