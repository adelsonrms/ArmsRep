VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sf_frmSelecionarItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub chkSelecionar_BeforeUpdate(Cancel As Integer)
    If Not pegaValor("MultiSelect") = "-1" Then
      '  CurrentDb.Execute "UPDATE [" & Me.RecordSource & "] SET selecionar = 0"
    End If
End Sub

Private Sub Form_Current()
    If Not pegaValor("MultiSelect") = "-1" Then
        'CurrentDb.Execute "UPDATE [" & Me.RecordSource & "] SET selecionar = 0"
        CurrentDb.Execute "UPDATE [" & Me.RecordSource & "] SET selecionar = 0 WHERE " & Me.rowID.ControlSource & "<>" & Me.rowID.value
        rowID.Requery
        'Me.Requery
        'chkSelecionar.Value = -1
        'CurrentDb.Execute "UPDATE [" & Me.RecordSource & "] SET selecionar = -1 where [" & txtCodigo.ControlSource & "]&[" & txtDescricao.ControlSource & "]='" & Me.txtCodigo.Value & Me.txtDescricao.Value & "'"
        'Me.Requery
    End If
End Sub

'''Private Sub chkSelecionar_Click()
'''    If Not getConfig("MultiSelect") = "-1" Then
'''        CurrentDb.Execute "UPDATE [" & Me.RecordSource & "] SET selecionar = 0 WHERE " & txtCodigo.ControlSource & "<>" & Me.txtCodigo.Value
'''    End If
'''End Sub

Private Sub Opção85_BeforeUpdate(Cancel As Integer)
    If Not pegaValor("MultiSelect") = "-1" Then
       ' CurrentDb.Execute "UPDATE [" & Me.RecordSource & "] SET selecionar = 0"
    End If
End Sub
