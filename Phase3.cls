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
    
    Dim salesData_ As SalesData
    Set salesData_ = Main.GetIndicator.GetRecord
    
    Dim ws As Worksheet
    Set ws = Worksheets("main")
    
    salesData_.SetLoss (Main.GetCabinet.DiscardItems)
    MessagePrinter.WriteLine salesData_.GetLoss & "個廃棄しました。"
    Main.GetIndicator.UpdateRecord salesData_
    
    salesData_.SetStock (salesData_.GetStock - salesData_.GetLoss)
    If (salesData_.GetStock < 0) Then
        salesData_.SetStock (0)
    End If
    
    Main.GetIndicator.UpdateRecord salesData_
    
    TablePrinter.PrintCurrentRecordToTable salesData_
    
    'TODO:ここに値引き後の在庫金額を書く。正値在庫＋値引き在庫
    
     MessagePrinter.WriteLine Strings.PLEASE_THE_ODER, 4
End Sub

Public Sub IPhase_ChangePhase()
    Main.SetCurrentPhase (PhaseNumber.LAST_4)
    MainSheet.Cells(MAIN_PHASE_PAIN.ROW, MAIN_PHASE_PAIN.Column) = Strings.PHASE_ODER
End Sub

