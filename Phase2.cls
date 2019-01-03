VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Phase2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'販売フェイズ

Implements IPhase

Public Sub IPhase_ExecutePhase()
    
    Dim record As SalesData
    Set record = Main.GetIndicator.GetRecord
    Dim cabinet_ As Cabinet
    Set cabinet_ = Main.GetCabinet
    
    cabinet_.DecideSales
    MessagePrinter.WriteLine record.GetSales & "個販売しました。"
    record.SetSales (cabinet_.GetSalesNum)
    record.SetSalesPrice (cabinet_.GetSalesPrice)
    record.SetPriceCut (cabinet_.GetPriceCutNum)
    record.SetStock (record.GetStock - record.GetSales)
    record.SetCustomerNum (cabinet_.GetCustomerNum)

    Main.GetIndicator.UpdateRecord sales_:=record.GetSales
    Main.GetIndicator.UpdateRecord priceCut_:=record.GetPriceCut
    Main.GetIndicator.UpdateRecord customerNum_:=record.GetCustomerNum
    Main.GetIndicator.UpdateRecord stock_:=record.GetStock
    
    If (record.GetStock <= 0) Then
        record.SetStock (0)
        MessagePrinter.WriteLine (Strings.OUT_OF_STOCK)
    End If
    
    TablePrinter.PrintCurrentRecordToTable record
   
End Sub

Public Sub IPhase_ChangePhase()
    Main.SetCurrentPhase (PhaseNumber.LOSS_3)
    MainSheet.Cells(MAIN_PHASE_PAIN.ROW, MAIN_PHASE_PAIN.COLUMN) = Strings.PHASE_LOSS
End Sub


