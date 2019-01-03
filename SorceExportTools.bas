Attribute VB_Name = "SorceExportTools"
Option Explicit

'�u�b�N���̃\�[�X�R�[�h���ꊇ�ŃG�N�X�|�[�g����c�[��
'�u�b�N���J����Ă��Ȃ��ꍇ�͌l�p�}�N���u�b�N�ipersonal.xlsb�j��ΏۂƂ���
'�u�b�N���J����Ă���ꍇ�͕\�����Ă���u�b�N��ΏۂƂ���

Sub ExportAll()
    Dim module As VBComponent       '���W���[��
    Dim moduleList As VBComponents  'VBA�v���W�F�N�g�̑S���W���[��
    Dim extension As String         '���W���[���̊g���q
    Dim sPath As String             '�����Ώۃu�b�N�̃p�X
    Dim sFilePath As String         '�G�N�X�|�[�g�t�@�C���p�X
    Dim TargetBook As Workbook      '�����Ώۃu�b�N�I�u�W�F�N�g
      
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
        Else '// �G�N�X�|�[�g�ΏۊO�̂��ߎ����[�v��
            GoTo CONTINUE
        End If
    
        sFilePath = sPath & "\src\" & module.Name & "." & extension
        Call module.Export(sFilePath)
        
        Debug.Print sFilePath
CONTINUE:
    Next
End Sub

