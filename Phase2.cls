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

'************************************************************************************
'販売フェイズ
'Store::ComingCustomerが商品の売れていく様子を再現
'************************************************************************************
Implements IPhase

Public Sub IPhase_ExecutePhase()
    
    Dim salesData_ As SalesData
    Set salesData_ = Main.GetIndicator.GetRecord
    Dim store_ As Store
    Set store_ = Main.GetStore
    
    store_.ComingCustomer
     
    salesData_.SetSales (store_.GetSalesNum)
    salesData_.SetSalesPrice (store_.GetSalesPrice)
    salesData_.SetPriceCut (store_.GetPriceCutNum)
    salesData_.SetStock (salesData_.GetStock - salesData_.GetSales)
    salesData_.SetCustomerNum (store_.GetCustomerNum)
       
    Main.GetIndicator.UpdateRecord salesData_
    MessagePrinter.WriteLine salesData_.GetSales & "個販売しました。 "
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


