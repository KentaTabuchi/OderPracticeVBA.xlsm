VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Indicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_row As Integer

Private Sub Class_Initialize()
    m_row = 2
End Sub


Public Function GetRow() As Integer
    GetRow = m_row
End Function

Public Function MoveOn() As SalesData
    If (m_row = 34) Then
        m_row = 34
    Else
        m_row = m_row + 1
    End If
    PrintCursor
    Set MoveOn = GetRecord
End Function

Public Function TurnBack() As SalesData
    If (m_row = 2) Then
        m_row = 2
    Else
        m_row = m_row - 1
    End If
    PrintCursor
    Set TurnBack = GetRecord
End Function

Public Function GetRecord(Optional row_ As Integer) As SalesData
    Dim ws As Worksheet
    Set ws = Worksheets("data")
    Dim salesData_ As SalesData
    Set salesData_ = New SalesData
    
    If row_ = 0 Then
        row_ = m_row
    End If
    salesData_.SetDate (ws.Cells(row_, DATA_COLUMN.DATE_2))
    salesData_.SetDelivery (ws.Cells(row_, DATA_COLUMN.delivery))
    salesData_.SetSales (ws.Cells(row_, DATA_COLUMN.sales))
    salesData_.SetLoss (ws.Cells(row_, DATA_COLUMN.loss))
    salesData_.SetPriceCut (ws.Cells(row_, DATA_COLUMN.priceCut))
    salesData_.SetCustomerNum (ws.Cells(row_, DATA_COLUMN.CUSTOMER_NUM))
    salesData_.SetStock (ws.Cells(row_, DATA_COLUMN.STOCK))
    
    Set GetRecord = salesData_
        
End Function
'dataシートのカレントレコードを指し示す→を表示する
Private Sub PrintCursor()
    Dim ws As Worksheet
    Set ws = Worksheets("data")
    ws.Columns(1).ClearContents
    ws.Cells(m_row, 1) = "→"
    
End Sub

'dataシートへの書き込み処理
Public Sub UpdateRecord(salesData_ As SalesData)
     
    Dim ws As Worksheet
    Set ws = Worksheets("data")
    ws.Cells(m_row, DATA_COLUMN.sales) = salesData_.GetSales
    ws.Cells(m_row, DATA_COLUMN.loss) = salesData_.GetLoss
    ws.Cells(m_row, DATA_COLUMN.priceCut) = salesData_.GetPriceCut
    ws.Cells(m_row, DATA_COLUMN.STOCK) = salesData_.GetStock
    ws.Cells(m_row, DATA_COLUMN.Oder) = salesData_.GetOder
    ws.Cells(m_row, DATA_COLUMN.CUSTOMER_NUM) = salesData_.GetCustomerNum
    
    With salesData_
    Debug.Print "GetSales:" & .GetSales, "GetLoss:" & .GetLoss, "GetPriceCut:" & .GetPriceCut, "GetStock:" & .GetStock, "GetOder:" & .GetOder
    End With
End Sub

'日付を引数にしてレコードを返す
Public Function FindRecordByDate(date_ As Date) As SalesData
    
    Dim ROW As Integer
    Dim i As Integer
    Dim ws As Worksheet: Set ws = Worksheets("data")
    
    For i = DATA_ROW.TABLE_TOP To DATA_ROW.TABLE_END
        If (ws.Cells(i, DATA_COLUMN.DATE_2) = date_) Then
            ROW = i
        End If
    Next i
    
    Dim salesData_ As SalesData
    Set salesData_ = GetRecord(ROW)
    Set FindRecordByDate = salesData_
    
End Function



