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
    
    Dim salesData_ As SalesData
    Set salesData_ = Main.GetIndicator.GetRecord
    Dim cabinet_ As Cabinet
    Set cabinet_ = Main.GetCabinet
    
    cabinet_.DecideSales
    
    MessagePrinter.WriteLine salesData_.GetSales & "個販売しました。"
    
    salesData_.SetSales (cabinet_.GetSalesNum)
    salesData_.SetSalesPrice (cabinet_.GetSalesPrice)
    salesData_.SetPriceCut (cabinet_.GetPriceCutNum)
    salesData_.SetStock (salesData_.GetStock - salesData_.GetSales)
    salesData_.SetCustomerNum (cabinet_.GetCustomerNum)

    Main.GetIndicator.UpdateRecord salesData_
    
    If (salesData_.GetStock <= 0) Then
        salesData_.SetStock (0)
        MessagePrinter.WriteLine (Strings.OUT_OF_STOCK)
    End If
    
    TablePrinter.PrintCurrentRecordToTable salesData_
   
End Sub

Public Sub IPhase_ChangePhase()
    Main.SetCurrentPhase (PhaseNumber.LOSS_3)
    MainSheet.Cells(MAIN_PHASE_PAIN.ROW, MAIN_PHASE_PAIN.Column) = Strings.PHASE_LOSS
End Sub


