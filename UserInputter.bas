Attribute VB_Name = "UserInputter"
Option Explicit

Private m_value As Integer

Public Function GetOderValue() As Integer

    Dim value_ As Variant
    value_ = MainSheet.Cells(20, 4)
    
    If IsNumeric(value_) Then
        m_value = Int(value_)
    Else
        MsgBox Strings.ERR_FAILED_INPUT
    End If

    GetOderValue = m_value
End Function

Public Function HasOdered() As Boolean
    
    Dim hasOdered_ As Boolean
    hasOdered_ = False
    
    If IsEmpty(MainSheet.Cells(20, 4)) Then
        MessagePrinter.WriteLine Strings.YET_ODER, 4
        
    Else
        hasOdered_ = True
    End If
   
    HasOdered = hasOdered_

End Function

Public Function Clear()
    MainSheet.Cells(20, 4).ClearContents
End Function
