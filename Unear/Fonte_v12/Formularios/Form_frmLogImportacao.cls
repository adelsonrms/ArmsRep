VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLogImportacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Activate()
    Stop
    Me.sfrmErrosDoLog.Requery
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: Form_frmLogImportacoes / Documento VBA
' Rotina....: Form_Load / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 22/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Inicio do tratamento do relatório de log de importação
'---------------------------------------------------------------------------------------
Private Sub Form_Load()
On Error GoTo TrataErro
Dim arrResumo           As Variant
Dim arrAnalitico        As Variant
Dim arrErros            As Variant
Dim arrDePara           As Variant

    Call Publicas.Inicializar_Globais
    Call Importacao.TratarRelatorioSaida(arrResumo, arrAnalitico, arrErros, arrDePara)
    
    With Me
        Call AuxForm.Atualizar_Campo_Versao(.lblVersao)
        .txtInicio.value = Importacao.ImportacaoInicio
        .txtFim.value = Importacao.ImportacaoFim

        If VBA.IsArray(arrResumo) Then
            Call PreencheListBox(arrResumo, Me.lstResumo)
        End If
        
        If VBA.IsArray(arrAnalitico) Then
            Call PreencheListBox(arrAnalitico, Me.lstAnalitico)
        End If
        
        If VBA.IsArray(arrErros) Then
            Call PreencheListBox(arrErros, Me.lstErros)
            .GuiaLogImportacao.Pages(2).Visible = True
            .GuiaLogImportacao.Pages(3).Visible = True
        Else
            .GuiaLogImportacao.Pages(2).Visible = False
            .GuiaLogImportacao.Pages(3).Visible = False
        End If
        If VBA.IsArray(arrDePara) Then
            Call PreencheListBox(arrDePara, Me.lstDePara)
            .GuiaLogImportacao.Pages(4).Visible = True
        Else
            .GuiaLogImportacao.Pages(4).Visible = False
        End If

        .txtTempo.value = VBA.Format(VBA.CDbl(Importacao.ImportacaoFim) - _
                                     VBA.CDbl(Importacao.ImportacaoInicio), "hh:MM:ss")
        .Requery
        .sfrmErrosDoLog.Requery
    End With
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Form_frmLogImportacao.Form_Load()", Erl)
End Sub
'---------------------------------------------------------------------------------------
' Modulo....: Form_frmLogImportacoes / Documento VBA
' Rotina....: PreencheListBox / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 22/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para criar e preencher o listBox
'---------------------------------------------------------------------------------------
Private Sub PreencheListBox(ByVal arrDados As Variant, ByRef lst As Access.ListBox)
On Error GoTo TrataErro
Dim valores         As String
Dim ContLinhas      As Long
Dim ContColunas     As Integer

    If VBA.IsArray(arrDados) Then
        With lst
            .ColumnCount = UBound(arrDados, 2) + 1
            .RowSourceType = "Value List"
            For ContLinhas = 0 To UBound(arrDados, 1) Step 1
                valores = VBA.vbNullString
                For ContColunas = 0 To UBound(arrDados, 2) Step 1
                    valores = valores & VBA.IIf(VBA.IsNumeric(arrDados(ContLinhas, ContColunas)), _
                                            VBA.Format(arrDados(ContLinhas, ContColunas), "hh:MM:ss"), _
                                            arrDados(ContLinhas, ContColunas)) & ";"
                Next ContColunas
                valores = VBA.Left(valores, VBA.Len(valores) - 1)
                .addItem valores
            Next ContLinhas
        End With
    End If
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Form_frmLogImportacao.PreencheListBox()", Erl)
End Sub


