VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Phase4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IPhase

Public Sub IPhase_ExecutePhase()
   
    If (UserInputter.HasOdered) Then
        Oder
        Main.GetIndicator.MoveOn
        TablePrinter.PrintLogRecordsToTable
        Main.SetCountDay (Main.GetCountDay + 1)
    End If
       
End Sub

Public Sub IPhase_ChangePhase()
    If (UserInputter.HasOdered) Then

        UserInputter.Clear
        Select Case Main.GetCountDay
        Case 0, 1:
            Main.SetCurrentPhase (PhaseNumber.LAST_4)
            MainSheet.Cells(MAIN_PHASE_PAIN.ROW, MAIN_PHASE_PAIN.Column) = Strings.PHASE_ODER
        Case Else:
            Main.SetCurrentPhase (PhaseNumber.START_0)
            MainSheet.Cells(MAIN_PHASE_PAIN.ROW, MAIN_PHASE_PAIN.Column) = Strings.PHASE_START
        End Select
    End If
End Sub

Private Sub Oder()
        Dim salesData_ As SalesData
        Set salesData_ = Main.GetIndicator.GetRecord
        salesData_.SetOder UserInputter.GetOderValue
        Main.GetIndicator.UpdateRecord salesData_
        MessagePrinter.WriteLine GetIndicator.GetRecord.GetOder & Strings.FINISH_ODER, 4
        
End Sub

