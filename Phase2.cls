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
    Debug.Print "phase:SELL2"
    
    Dim record As SalesData
    Set record = Main.GetIndicator.GetRecord
    Dim cabinet_ As Cabinet
    Set cabinet_ = Main.GetCabinet
    
    record.SetSales (cabinet_.DecideSales)
    record.SetSalesPrice (cabinet_.GetSalesPrice)
    record.SetPriceCut (cabinet_.GetPriceCutNum)
    record.SetStock (record.GetStock - record.GetSales)
    record.SetCustomerNum (cabinet_.GetCustomerNum)

    Main.GetIndicator.UpdateRecord sales_:=record.GetSales
    Main.GetIndicator.UpdateRecord priceCut_:=record.GetPriceCut
    Main.GetIndicator.UpdateRecord customerNum_:=record.GetCustomerNum
    
    If (record.GetStock <= 0) Then
        record.SetStock (0)
        Message.WriteLine (Strings.str_OutOfStock)
    End If
    
    Dim ws1 As Worksheet
    Set ws1 = Worksheets("main")
    ws1.Cells(MAIN_ROW.sales, MAIN_COLUMN.number) = record.GetSales
    ws1.Cells(MAIN_ROW.sales, MAIN_COLUMN.Price) = record.GetSalesPrice
   
    Dim ws2 As Worksheet
    Set ws2 = Worksheets("data")
    ws2.Cells(Main.GetIndicator.GetRow, DATA_COLUMN.sales) = record.GetSales
End Sub

Public Sub IPhase_ChangePhase()
    Main.SetCurrentPhase (PhaseNumber.LOSS_3)
End Sub


