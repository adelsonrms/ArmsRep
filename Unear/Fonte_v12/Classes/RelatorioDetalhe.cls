VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RelatorioDetalhe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============PROPRIEDADES=============
Private m_Posicao           As Integer
Private m_ID                As Integer
Private m_Referencia        As String
Private m_NomeRelatorio     As String
Private m_NomeArquivo       As String
Private m_FileNameImport    As String
Private m_Tabela            As Tabelas
Private m_ParentID          As Integer
Private m_Identificacao     As String
Private m_Regex             As String
Private m_TipoRelatorio     As String
Private m_Visivel           As Boolean
Private m_NomeExibicao      As String
Private m_Vinculada         As Boolean
Private m_eTipoLeitura As TipoLeitura

Public Property Get Posicao() As Integer
    Posicao = m_Posicao
End Property
Public Property Let Posicao(ByVal valor As Integer)
    m_Posicao = valor
End Property

Public Property Get ID() As Integer
    ID = m_ID
End Property
Public Property Let ID(ByVal valor As Integer)
    m_ID = valor
End Property

Public Property Get Referencia() As String
    Referencia = m_Referencia
End Property
Public Property Let Referencia(ByVal valor As String)
    m_Referencia = valor
End Property

Public Property Get NomeRelatorio() As String
    NomeRelatorio = m_NomeRelatorio
End Property
Public Property Let NomeRelatorio(ByVal valor As String)
    m_NomeRelatorio = valor
End Property

Public Property Get NomeArquivo() As String
    NomeArquivo = m_NomeArquivo
End Property


Public Property Get FileNameImport() As String
    FileNameImport = m_FileNameImport
End Property
Public Property Let FileNameImport(ByVal valor As String)
    m_FileNameImport = valor
End Property

Public Property Let NomeArquivo(ByVal valor As String)
    m_NomeArquivo = valor
End Property

Public Property Get Tabela() As Tabelas
    Set Tabela = m_Tabela
End Property
Public Property Set Tabela(ByVal valor As Tabelas)
    Set m_Tabela = valor
End Property

Public Property Get ParentID() As String
    ParentID = m_ParentID
End Property
Public Property Let ParentID(ByVal valor As String)
    m_ParentID = valor
End Property

Public Property Get Identificacao() As String
    Identificacao = m_Identificacao
End Property
Public Property Let Identificacao(ByVal valor As String)
    m_Identificacao = valor
End Property

Public Property Get Regex() As String
    Regex = m_Regex
End Property
Public Property Let Regex(ByVal valor As String)
    m_Regex = valor
End Property

Public Property Get TipoRelatorio() As String
    TipoRelatorio = m_TipoRelatorio
End Property
Public Property Let TipoRelatorio(ByVal valor As String)
    m_TipoRelatorio = valor
End Property

Public Property Get Visivel() As Boolean
    Visivel = m_Visivel
End Property
Public Property Let Visivel(ByVal valor As Boolean)
    m_Visivel = valor
End Property

Public Property Get NomeExibicao() As String
    NomeExibicao = m_NomeExibicao
End Property
Public Property Let NomeExibicao(ByVal valor As String)
    m_NomeExibicao = valor
End Property

Public Property Get Vinculada() As String
    Vinculada = m_Vinculada
End Property
Public Property Let Vinculada(ByVal valor As String)
    m_Vinculada = valor
End Property

'=============SUB-ROTINAS=============
Public Sub SetPropriedades(ByVal campo As String, ByVal valor As Variant)
On Error GoTo TrataErro

    If VBA.IsNull(valor) Then valor = VBA.vbNullString
    With Me
        Select Case campo
            Case "Posicao"
                .Posicao = valor
            Case "RelID"
                If Not VBA.IsNumeric(valor) Then valor = 0
                .ID = VBA.CInt(valor)
            Case "Referencia"
                .Referencia = VBA.CStr(valor)
            Case "NomeRelatorio"
                .NomeRelatorio = VBA.CStr(valor)
            Case "NomeArquivo"
                .NomeArquivo = VBA.CStr(valor)
            Case "Tabela"
                Set .Tabela = New Tabelas
                .Tabela.Name = VBA.CStr(valor)
            Case "ParentID"
                If Not VBA.IsNumeric(valor) Then valor = 0
                .ParentID = VBA.CInt(valor)
            Case "Identificacao"
                .Identificacao = VBA.CStr(valor)
            Case "Regex"
                .Regex = VBA.CStr(valor)
            Case "TipoRelatorio"
                .TipoRelatorio = VBA.CStr(valor)
            Case "Visivel"
                If valor = VBA.vbNullString Then valor = False
                .Visivel = VBA.CBool(valor)
            Case "NomeExibicao"
                .NomeExibicao = VBA.CStr(valor)
            Case "Vinculada"
                .Vinculada = VBA.CBool(valor)
            Case "TipoLeitura"
                .TipoLeitura = valor
        End Select
    End With
Exit Sub
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "RelatorioColunas.SetPropriedades()", Erl)
End Sub


Public Property Get TipoLeitura() As TipoLeitura

    TipoLeitura = m_eTipoLeitura

End Property

Public Property Let TipoLeitura(ByVal eTipoLeitura As TipoLeitura)

    m_eTipoLeitura = eTipoLeitura

End Property
