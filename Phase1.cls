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
    
    Dim salesData_ As SalesData
    Set salesData_ = Main.GetIndicator.GetRecord

    salesData_.SetStock (salesData_.GetStock + salesData_.GetDelivery)
    Main.GetIndicator.UpdateRecord salesData_
    Main.GetCabinet.DisplayItems (Main.GetIndicator.GetRecord.GetDelivery)
    TablePrinter.PrintCurrentRecordToTable salesData_
    
End Sub

Public Sub IPhase_ChangePhase()
    Main.SetCurrentPhase (PhaseNumber.SELL_2)
    MainSheet.Cells(MAIN_PHASE_PAIN.ROW, MAIN_PHASE_PAIN.Column) = Strings.PHASE_SELL
End Sub
