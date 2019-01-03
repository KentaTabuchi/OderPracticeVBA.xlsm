Attribute VB_Name = "TablePrinter"
Option Explicit

'メインシートのテーブルにデータシートのレコードを表示するためのモジュール。

Private Type TableUnit
    num_column As Integer
    price_column As Integer
    date_row As Integer
    carryOver_row As Integer
    priceCut_row As Integer
    delivery_row As Integer
    sales_row As Integer
    loss_row As Integer
    currentStock_row As Integer
End Type

Private m_TableUnits(7) As TableUnit
'７日分の過去レコードをテーブルに表示
Public Sub PrintLogRecordsToTable()

    SetTableMap

    Dim salesData_ As SalesData
    
    Dim indicator_ As Indicator
    Set indicator_ = Main.GetIndicator
    
    Dim row_ As Integer
    row_ = indicator_.GetRow
    
    Dim i As Integer
    Dim ws As Worksheet
    Set ws = Worksheets("main")
    
    For i = 1 To 6
        ClearTableUnit (i)
    Next i
    For i = 1 To 6
        If (row_ - i < 2) Then Exit For
        Set salesData_ = indicator_.GetRecord(row_ - i + 1)
        With m_TableUnits(i)
            ws.Cells(.date_row, .num_column) = salesData_.GetDate
            ws.Cells(.carryOver_row, .num_column) = salesData_.GetStock
            ws.Cells(.delivery_row, .num_column) = salesData_.GetDelivery
            ws.Cells(.loss_row, .num_column) = salesData_.GetLoss
            ws.Cells(.priceCut_row, .num_column) = salesData_.GetPriceCut
            ws.Cells(.sales_row, .num_column) = salesData_.GetSales
        End With
     
    Next i
       
End Sub

'当日の速報をテーブルに表示していくメソッド
'基本的にdataシートの当日の行を毎回全部更新しているだけだが、繰り越し在庫は
'前日分の確定在庫と同値になるので一行上から引っ張るようにしている。（二つ目のwith部分）
Public Sub PrintCurrentRecordToTable(salesData_ As SalesData)
    
    SetTableMap
   
    Dim indicator_ As Indicator
    Set indicator_ = Main.GetIndicator
    
    Dim row_ As Integer
    row_ = indicator_.GetRow
    
    With m_TableUnits(0)
        MainSheet.Cells(.date_row, .num_column) = salesData_.GetDate
        MainSheet.Cells(.currentStock_row, .num_column) = salesData_.GetStock
        MainSheet.Cells(.delivery_row, .num_column) = salesData_.GetDelivery
        MainSheet.Cells(.loss_row, .num_column) = salesData_.GetLoss
        MainSheet.Cells(.priceCut_row, .num_column) = salesData_.GetPriceCut
        MainSheet.Cells(.sales_row, .num_column) = salesData_.GetSales
    End With
        
    Dim previousDaySalesData_ As SalesData
    Set previousDaySalesData_ = indicator_.GetRecord(row_ - 1)
    
    With m_TableUnits(0)
        MainSheet.Cells(.carryOver_row, .num_column) = previousDaySalesData_.GetStock
    End With
    
End Sub
Private Sub SetTableMap()
    
    Dim i As Integer
    For i = 0 To 7
     With m_TableUnits(i)
     .num_column = MAIN_COLUMN.TABLE_LEFT_EDGE + (i * 2)
     .price_column = .num_column + 1
     .date_row = MAIN_ROW.date_
     .carryOver_row = MAIN_ROW.CARRY_OVER_STOCK
     .delivery_row = MAIN_ROW.delivery
     .loss_row = MAIN_ROW.loss
     .priceCut_row = MAIN_ROW.PRICE_CUT
     .currentStock_row = MAIN_ROW.CURRENT_STOCK
     .sales_row = MAIN_ROW.sales
     End With
    Next i
    
End Sub
'mainシートのテーブルの数字を消す。
Private Sub ClearTableUnit(index As Integer)

    SetTableMap

    Dim ws As Worksheet
    Set ws = Worksheets("main")

        With m_TableUnits(index)
            ws.Cells(.date_row, .num_column).ClearContents
            ws.Cells(.carryOver_row, .num_column).ClearContents
            ws.Cells(.delivery_row, .num_column).ClearContents
            ws.Cells(.loss_row, .num_column).ClearContents
            ws.Cells(.priceCut_row, .num_column).ClearContents
            ws.Cells(.sales_row, .num_column).ClearContents
            
            ws.Cells(.date_row, .price_column).ClearContents
            ws.Cells(.carryOver_row, .price_column).ClearContents
            ws.Cells(.delivery_row, .price_column).ClearContents
            ws.Cells(.loss_row, .price_column).ClearContents
            ws.Cells(.priceCut_row, .price_column).ClearContents
            ws.Cells(.sales_row, .price_column).ClearContents
        End With

End Sub

