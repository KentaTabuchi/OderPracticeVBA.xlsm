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
    
    Dim record As SalesData
    Set record = Main.GetIndicator.GetRecord

    record.SetStock (record.GetStock + record.GetDelivery)
    Main.GetIndicator.UpdateRecord stock_:=record.GetStock
    Main.GetCabinet.DisplayItems (Main.GetIndicator.GetRecord.GetDelivery)
    TablePrinter.PrintCurrentRecordToTable record
    
End Sub

Public Sub IPhase_ChangePhase()
    Main.SetCurrentPhase (PhaseNumber.SELL_2)
    MainSheet.Cells(MAIN_PHASE_PAIN.ROW, MAIN_PHASE_PAIN.COLUMN) = Strings.PHASE_SELL
End Sub
