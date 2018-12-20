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
    Debug.Print "phase:LAST_4"
    
    'TODO:ここで日付変更になるので表の日付を一列づつずらす処理を書く
     
    If (IsOder) Then
        Oder
        Main.GetIndicator.MoveOn
        Scroll
    End If
       
End Sub

Public Sub IPhase_ChangePhase()
    If (IsOder) Then
        Dim ws As Worksheet
        Set ws = Worksheets("main")
        ws.Cells(20, 4).ClearContents
        Main.SetCurrentPhase (PhaseNumber.START_0)
    End If
End Sub

Private Function IsOder() As Boolean
    IsOder = False
    
    Dim ws As Worksheet
    Set ws = Worksheets("main")
    If (ws.Cells(20, 4) <> "") Then
        IsOder = True
    Else
        Message.WriteLine Strings.str_YetOder, 4
    End If
   
End Function

Private Sub Oder()
        Dim ws As Worksheet
        Set ws = Worksheets("main")
        Main.GetIndicator.UpdateRecord oder_:=ws.Cells(20, 4)
        Message.WriteLine GetIndicator.GetRecord.GetOder & Strings.str_FinishOder, 4
        
End Sub
Private Sub Scroll()
    Dim ws As Worksheet
    Set ws = Worksheets("main")
    Dim i As Integer
    For i = 21 To 6 Step -3
        Range(Cells(2, i), Cells(8, i + 2)).Copy Cells(2, i + 3)
        Range(Cells(15, i), Cells(100, i + 2)).Copy Cells(15, i + 3)
    Next i
End Sub
