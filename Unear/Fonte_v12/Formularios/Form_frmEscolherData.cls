VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmEscolherData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'********************PROPRIEDADES********************

Private m_DataCredito As Date

Public Property Get DataCredito() As Date
    DataCredito = m_DataCredito
End Property
Public Property Let DataCredito(ByVal dtValor As Date)
    m_DataCredito = dtValor
End Property

Private Sub cmdHoje_Click()
    txtData.value = cmdHoje.tag
End Sub

Private Sub Form_Load()
On Error GoTo TrataErro
Dim arrArgs     As Variant
Dim strFormato      As String

    With Me
        If Not VBA.IsNull(.OpenArgs) Then
            arrArgs = VBA.Split(.OpenArgs, "|")
        End If
        If VBA.IsArray(arrArgs) Then
            If Not arrArgs(1) = VBA.Space(1) Then
                .lblData.Caption = VBA.CStr(arrArgs(1))
            Else
                .lblData.Caption = "Digite a Data de Referência"
            End If
            With .txtData
                If VBA.IsDate(arrArgs(0)) Then
                    strFormato = VBA.Replace(VBA.Replace(VBA.Replace(VBA.Replace(arrArgs(2), "/0000", "/yyyy"), "00/y", "mm/y"), "00/m", "dd/m"), "!", "")
                    If VBA.CDate(arrArgs(0)) > 0 Then .value = VBA.Format(arrArgs(0), strFormato)
                End If
                .InputMask = arrArgs(2)
                .SetFocus
            End With
        Else 'VALORES PADROES
            .txtData.InputMask = "!00/00/0000;0;_"
            .lblData.Caption = "Digite a Data de Referência"
        End If
        cmdHoje.tag = VBA.Format(Date, strFormato)
        cmdHoje.Caption = "Hoje > " & cmdHoje.tag
    End With
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "frmEscolherData.Form_Load()")
End Sub

Private Sub btnOK_Click()
On Error GoTo TrataErro
    With Me.txtData
        .SetFocus
        If .Text = VBA.vbNullString Then
            Call AuxMensagens.MessageBoxMaster("F010")
        ElseIf Not VBA.IsDate(.Text) Then
            Call AuxMensagens.MessageBoxMaster("F011")
        Else
            DataCredito = VBA.CDate(.Text)
            Call salvaValor("DataSelecionada", DataCredito)
            Call salvaValor("periodo", VBA.Format(DataCredito, "yyyymm"))
            Me.Visible = False
            btnCancelar_Click
        End If
    End With
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "frmEscolherData.btnOK_Click()")
End Sub

Private Sub btnCancelar_Click()
On Error GoTo TrataErro
    Call Access.DoCmd.Close(Access.AcObjectType.acForm, Me.Name)
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "frmEscolherData.btnCancelar_Click()")
End Sub

