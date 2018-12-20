Attribute VB_Name = "Message"
Option Explicit

'Mainシートのメッセージ欄を操作するコントローラー

Public Sub WriteLine(Message As String, Optional line As Integer = 1)
    Dim ws As Worksheet
    Set ws = Worksheets("main")
    

    ws.Cells(14 + line, 2) = Message
End Sub
Public Sub ClearAll()
    Dim ws As Worksheet
    Set ws = Worksheets("main")
    
    ws.Cells(15, 2).ClearContents
    ws.Cells(16, 2).ClearContents
    ws.Cells(17, 2).ClearContents
    ws.Cells(18, 2).ClearContents
    
End Sub
