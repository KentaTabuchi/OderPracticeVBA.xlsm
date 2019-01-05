Attribute VB_Name = "Main"
Option Explicit

Private m_phases(0 To 4) As IPhase
Private m_isInitialize As Boolean
Private m_indicator As Indicator
Private m_currentPhase As PhaseNumber
Private m_store As Store
Private m_countDay As Integer

Public Sub OnAdvanceTheTimeButtonClick()

    If m_isInitialize = False Then
        Initialize
        m_isInitialize = True
    End If
    
    Select Case m_countDay
    Case 0:
        MessagePrinter.WriteLine Strings.INPUT_FIRST_ODER
        ExecutePhase (PhaseNumber.LAST_4)
    Case 1:
        MessagePrinter.WriteLine Strings.INPUT_SECOND_ODER
        ExecutePhase (PhaseNumber.LAST_4)
    Case Else
        ExecutePhase (m_currentPhase)
    End Select
End Sub

Public Sub ExecutePhase(phase_ As PhaseNumber)
    m_phases(phase_).ExecutePhase
    m_phases(phase_).ChangePhase
End Sub

Private Sub Initialize()
    
    DevelopmentTools.ClearImmediate
    
    Set m_indicator = New Indicator
    Set m_store = New Store
    m_currentPhase = PhaseNumber.START_0
    
    Set m_phases(0) = New Phase0
    Set m_phases(1) = New Phase1
    Set m_phases(2) = New Phase2
    Set m_phases(3) = New Phase3
    Set m_phases(4) = New Phase4
    
    Dim ws As Worksheet
    Set ws = Worksheets("main")
    ws.Cells(MAIN_ROW.date_, 5).ClearContents
    ws.Cells(MAIN_ROW.CARRY_OVER_STOCK, 5).ClearContents
    ws.Cells(MAIN_ROW.delivery, 5).ClearContents
    ws.Cells(MAIN_ROW.sales, 5).ClearContents
    ws.Cells(MAIN_ROW.loss, 5).ClearContents
    
    MessagePrinter.ClearAll
    ClearTable
    MainSheet.Cells(MAIN_PHASE_PAIN.ROW, MAIN_PHASE_PAIN.Column) = Empty
    
End Sub
Public Function GetIndicator() As Indicator
    Set GetIndicator = m_indicator
End Function
Public Function GetCurrentPhase() As PhaseNumber
    GetCurrentPhase = m_currentPhase
End Function
Public Sub SetCurrentPhase(phase_ As PhaseNumber)
    m_currentPhase = phase_
End Sub
Public Function GetStore() As Store
    Set GetStore = m_store
End Function
Public Function GetCountDay() As Integer
    GetCountDay = m_countDay
End Function
Public Sub SetCountDay(value_ As Integer)
    m_countDay = value_
End Sub


Private Sub ClearTable()
    Dim ws As Worksheet
    Set ws = Worksheets("main")
    
    ws.Range(Cells(2, 6), Cells(7, 26)).ClearContents
    ws.Range(Cells(16, 6), Cells(100, 26)).ClearContents
    
End Sub

