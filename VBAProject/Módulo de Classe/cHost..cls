Private WithEvents appHost As Application
Private Editor As VBEditor

Private Sub appHost_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    Debug.Print "Iniciando backup de : " & Wb.Name
    Call Editor.BackCode
    Debug.Print "Backup Finalizado : " & Wb.Name
End Sub

Private Sub appHost_WorkbookOpen(ByVal Wb As Workbook)
    Debug.Print "Addin inicializado para o projeto : " & Wb.Name
End Sub

Private Sub Class_Initialize()
    Set Editor = New VBEditor
    Editor.PastaBackup = "P:\Store\Cloud\GitHub\Repositorios\ArmsRep"
    Set appHost = Application
End Sub
