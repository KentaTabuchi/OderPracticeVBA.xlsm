Attribute VB_Name = "DevelopmentTools"
Option Explicit

'ブック内のソースコードを一括でエクスポートするツール
'ブックが開かれていない場合は個人用マクロブック（personal.xlsb）を対象とする
'ブックが開かれている場合は表示しているブックを対象とする

Sub ExportAll()
    Dim module As VBComponent       'モジュール
    Dim moduleList As VBComponents  'VBAプロジェクトの全モジュール
    Dim extension As String         'モジュールの拡張子
    Dim sPath As String             '処理対象ブックのパス
    Dim sFilePath As String         'エクスポートファイルパス
    Dim TargetBook As Workbook      '処理対象ブックオブジェクト
      
    If (Workbooks.Count = 1) Then
        Set TargetBook = ThisWorkbook
    Else
        Set TargetBook = ActiveWorkbook
    End If
       
    sPath = TargetBook.Path
  
    Set moduleList = TargetBook.VBProject.VBComponents
  
    For Each module In moduleList
       
        If (module.Type = vbext_ct_ClassModule) Then
            extension = "cls"
        ElseIf (module.Type = vbext_ct_MSForm) Then
            extension = "frm"
        ElseIf (module.Type = vbext_ct_StdModule) Then
            extension = "bas"
        Else
            GoTo CONTINUE
        End If
    
        sFilePath = sPath & "\src\" & module.Name & "." & extension
        Call module.Export(sFilePath)
        
        Debug.Print sFilePath
CONTINUE:
    Next
End Sub

'******************************
' イミディエイトウィンドウのクリア
'******************************
Public Sub ClearImmediate()
 
    Dim wd      As VBIDE.Window
    Dim wdwk    As VBIDE.Window
     
    '*** イミディエイトウィンドウの取得
    Set wd = Application.VBE.Windows("イミディエイト")
    If Not wd.Visible Then Exit Sub     '非表示だったら抜け
     
    '*** イミディエイトウィンドウのクリア
    wd.SetFocus
    SendKeys "^a", False
    SendKeys "{Del}", False
 
End Sub

