VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Phase0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IPhase

Public Sub IPhase_ExecutePhase()

    MessagePrinter.ClearAll
       
    Dim record As SalesData
    Set record = Main.GetIndicator.GetRecord
    
    TablePrinter.PrintCurrentRecordToTable record
  
End Sub

Public Sub IPhase_ChangePhase()
    Main.SetCurrentPhase (PhaseNumber.DELIVERY_1)
    MainSheet.Cells(MAIN_PHASE_PAIN.ROW, MAIN_PHASE_PAIN.COLUMN) = Strings.PHASE_DELIVERY
End Sub


