Attribute VB_Name = "UnitTest"
Option Explicit

'�P�̃e�X�g�p���W���[��

Public Sub TestMain()

    Dim indicator_ As Indicator
    Set indicator_ = New Indicator

    Dim salesData_ As SalesData
    Set salesData_ = indicator_.FindRecordByDate(#12/2/2018#)
    
    With salesData_
    
    Debug.Print .GetDate, .GetDelivery, .GetLoss
    
    End With
    
    
End Sub


