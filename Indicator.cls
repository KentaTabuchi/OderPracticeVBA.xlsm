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
Private m_salesData As SalesData

Private Sub Class_Initialize()
    m_row = 2
    Set m_salesData = New SalesData
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
    
    If row_ = 0 Then
        row_ = m_row
    End If
    m_salesData.SetDate (ws.Cells(row_, DATA_COLUMN.DATE_2))
    m_salesData.SetDelivery (ws.Cells(row_, DATA_COLUMN.delivery))
    m_salesData.SetSales (ws.Cells(row_, DATA_COLUMN.sales))
    m_salesData.SetLoss (ws.Cells(row_, DATA_COLUMN.loss))
    m_salesData.SetPriceCut (ws.Cells(row_, DATA_COLUMN.priceCut))
    m_salesData.SetCustomerNum (ws.Cells(row_, DATA_COLUMN.CUSTOMER_NUM))
    m_salesData.SetStock (ws.Cells(row_, DATA_COLUMN.STOCK))
    
    Set GetRecord = m_salesData
        
End Function
'dataシートのカレントレコードを指し示す→を表示する
Private Sub PrintCursor()
    Dim ws As Worksheet
    Set ws = Worksheets("data")
    ws.Columns(1).ClearContents
    ws.Cells(m_row, 1) = "→"
    
End Sub

'dataシートへの書き込み処理
Public Sub UpdateRecord(Optional sales_ As Integer, Optional loss_ As Integer, Optional stock_ As Integer, Optional oder_ As Integer, Optional priceCut_ As Integer, Optional customerNum_ As Integer)
    
    If (sales_ <> Empty) Then m_salesData.SetSales (sales_)
    If (loss_ <> Empty) Then m_salesData.SetLoss (loss_)
    If (stock_ <> Empty) Then m_salesData.SetStock (stock_)
    If (oder_ <> Empty) Then m_salesData.SetOder (oder_)
    If (priceCut_ <> Empty) Then m_salesData.SetPriceCut (priceCut_)
    If (customerNum_ <> Empty) Then m_salesData.SetCustomerNum (customerNum_)
    
     
    Dim ws As Worksheet
    Set ws = Worksheets("data")
    ws.Cells(m_row, DATA_COLUMN.sales) = m_salesData.GetSales
    ws.Cells(m_row, DATA_COLUMN.loss) = m_salesData.GetLoss
    ws.Cells(m_row, DATA_COLUMN.priceCut) = m_salesData.GetPriceCut
    ws.Cells(m_row, DATA_COLUMN.STOCK) = m_salesData.GetStock
    ws.Cells(m_row, DATA_COLUMN.Oder) = m_salesData.GetOder
    ws.Cells(m_row, DATA_COLUMN.CUSTOMER_NUM) = m_salesData.GetCustomerNum
    
    With m_salesData
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



