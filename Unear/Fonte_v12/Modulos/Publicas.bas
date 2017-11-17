Attribute VB_Name = "Publicas"
Option Compare Database
Option Explicit

Public bStop As Boolean
Public errNumber As Long
Public errDescricao As String
Private m_sChaveUsuario  As String
''''---------------------------------------------------------------------------------------
'''' Modulo....: Publicas / Módulo
'''' Rotina....: ChaveUsuario / Property
'''' Autor.....: Jefferson Dantas
'''' Contato...: jefferson@tecnun.com.br
'''' Data......: 14/08/2013
'''' Empresa...: Tecnun Tecnologia em Informática
'''' Descrição.: Propriedade readonly que cria a chave do usuario baseado no username e
''''             computername usando a ferramenta no momento, além disso a rotina remove
''''             qualquer caracter invalido.
''''---------------------------------------------------------------------------------------
'''Public Property Get ChaveUsuario() As String
'''Dim contador            As Long
'''Dim ChaveAux            As String
'''Dim ChaveResultado      As String
'''
'''    ChaveAux = VBA.UCase(VBA.Environ("Username") & "_" & VBA.Environ("ComputerName"))
'''    For contador = 1 To VBA.Len(ChaveAux) Step 1
'''        If IsLinhaMatch(VBA.Mid(ChaveAux, contador, 1), "([A-Z]|[a-z]|[0-9])|[_]") Then
'''            ChaveResultado = ChaveResultado & VBA.Mid(ChaveAux, contador, 1)
'''        End If
'''    Next contador
'''    ChaveUsuario = ChaveResultado
'''End Property

Public Property Get ChaveUsuario() As String
    m_sChaveUsuario = VBA.Environ("Username") & "_" & VBA.Environ("ComputerName")
    m_sChaveUsuario = VBA.Replace(m_sChaveUsuario, " ", "")
    ChaveUsuario = UCase(m_sChaveUsuario)
End Property



'********************SUB-ROTINAS********************
'---------------------------------------------------------------------------------------
' Modulo    : Publicas / Módulo
' Rotina    : RemoverObjetosMemoria() / Sub
' Autor     : Jefferson
' Data      : 07/11/2012 - 16:42
' Proposta  : Rotina para remover os objetos da memoria
'---------------------------------------------------------------------------------------
Public Sub RemoverObjetosMemoria(ParamArray Objetos() As Variant)
On Error Resume Next 'Resume next necessario em caso de erro
Dim contador    As Integer
    For contador = 0 To UBound(Objetos) Step 1
        If VBA.TypeName(Objetos(contador)) = "Variant()" Then
            Objetos(contador) = Empty
        Else
            Set Objetos(contador) = Nothing
        End If
    Next contador
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: Publicas / Módulo
' Rotina....: Inicializar_Globais / Sub
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 15/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para inicializar as variaveis globais podendo forçar a inicialização
'---------------------------------------------------------------------------------------
Public Sub Inicializar_Globais(Optional ByVal ForcaInicializacao As Boolean = False)
On Error GoTo TratarErro
    If VariaveisEConstantes.Conexao Is Nothing Or ForcaInicializacao Then
        Set VariaveisEConstantes.Conexao = New ConexaoDB
    End If
    If VariaveisEConstantes.Relatorio Is Nothing Or ForcaInicializacao Then
        Set VariaveisEConstantes.Relatorio = New Relatorios
        Set VariaveisEConstantes.Relatorio.CDB_Apoio.ConexaoBanco = VariaveisEConstantes.Connection_Cliente
    End If
    '--------------------------------------------------------------------
    'Causa chamada recursiva
    'Call Carregar_Versao
    '--------------------------------------------------------------------
On Error GoTo 0
Exit Sub
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Publicas.Inicializar_Globais", Erl)
End Sub

'Para compatibilidade com versões antigas
Sub VerificaGlobais()
    Inicializar_Globais
End Sub


Public Function Carregar_Versao()
On Error GoTo TratarErro
Dim VersaoAux       As String
    VersaoAux = CurrentDb.OpenRecordset("Pegar_Versao_Atual")(0)
    If Not VersaoAux = VBA.vbNullString Then
        If VBA.Right(VersaoAux, 1) = 0 Then
            VersaoAux = VersaoAux & " - Desenvolvimento"
        ElseIf VBA.Right(VersaoAux, 1) = 1 Then
            VersaoAux = VersaoAux & " - Homologação"
        ElseIf VBA.Right(VersaoAux, 1) = 2 Then
            VersaoAux = VersaoAux & " - Produção"
        End If
        VersaoAux = "v" & VersaoAux
        VariaveisEConstantes.versao = VersaoAux
        Carregar_Versao = VersaoAux
    End If
On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Publicas.Carregar_Versao", Erl)
End Function

Function PegarVersao() As eVersao
    Select Case VBA.Split(VBA.Split(Carregar_Versao(), "-")(0), ".")(2)
    Case 0: PegarVersao = Desenvolvimento
    Case 1: PegarVersao = Homogolacao
    Case 2: PegarVersao = Producao
    End Select
End Function


'''Public Sub Carregar_Versao()
'''On Error GoTo TratarErro
'''Dim VersaoAux       As String
'''    If Conexao Is Nothing Then Call Publicas.Inicializar_Globais
'''    'VersaoAux = Conexao.PegarString("|", "Pegar_Versao_Atual")
'''    VersaoAux = CurrentDb.OpenRecordset("Pegar_Versao_Atual")(0)  'Conexao.PegarString("|", "Pegar_Versao_Atual")
'''    If Not VersaoAux = VBA.vbNullString Then
'''        If VBA.Right(VersaoAux, 1) = 0 Then
'''            VersaoAux = VersaoAux & " - Desenvolvimento"
'''        ElseIf VBA.Right(VersaoAux, 1) = 1 Then
'''            VersaoAux = VersaoAux & " - Homologação"
'''        ElseIf VBA.Right(VersaoAux, 1) = 2 Then
'''            VersaoAux = VersaoAux & " - Produção"
'''        End If
'''        VersaoAux = "v" & VersaoAux
'''        VariaveisEConstantes.versao = VersaoAux
'''    End If
'''On Error GoTo 0
'''Exit Sub
'''TratarErro:
'''    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Publicas.Carregar_Versao", Erl)
'''End Sub

'********************FUNCOES********************
'---------------------------------------------------------------------------------------
' Modulo....: Publicas / Módulo
' Rotina....: ChaveExists / Function
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 19/12/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina criada para saber se uma chave existe dentro de um dicionario,
'             usando a rotina de islinhamatch para auxiliar
'---------------------------------------------------------------------------------------
Public Function ChaveExists(ByRef ChaveEncontrada As String, ByRef dic As Object, _
                            ByVal chave As String) As Boolean
On Error GoTo TratarErro
Dim dicChave            As Variant
Dim ChaveAux            As String
Dim Resultado           As Boolean

    If dic.count > 0 Then
        For Each dicChave In dic
            ChaveAux = PegarTexto_Regex(RemoverAcentos(chave, False), _
                                                 RemoverAcentos(dicChave, False))
            If Not ChaveAux = VBA.vbNullString Then
                ChaveEncontrada = dicChave
                Resultado = True
                Exit For
            End If
        Next dicChave
    End If
    ChaveExists = Resultado
On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Publicas.ChaveExists", Erl)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: Publicas / Módulo
' Rotina....: EnderecoValido / Function
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 14/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina que verifica se um determinado endereço é valido ou não
'---------------------------------------------------------------------------------------
Public Function EnderecoValido(ByVal strEndereco As String, ByVal endereco As tipoEndereco) As Boolean
On Error GoTo TrataErro
Dim fso             As Object
Dim fsoFile         As Object
Dim fsoFolder       As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    Select Case endereco
        Case tipoEndereco.Pastas
            Set fsoFolder = fso.GetFolder(strEndereco)
        Case tipoEndereco.arquivos
            Set fsoFile = fso.GetFile(strEndereco)
    End Select

    EnderecoValido = True
    Call Publicas.RemoverObjetosMemoria(fsoFile, fsoFolder, fso)
Exit Function
TrataErro:
'Não é necessário logar esse tipo de erro.
    EnderecoValido = False
    Call Publicas.RemoverObjetosMemoria(fsoFile, fsoFolder, fso)
End Function

'---------------------------------------------------------------------------------------
' Modulo....: Publicas / Módulo
' Rotina....: GerarChaves / Function
' Autor.....: Jefferson Dantas
' Contato...: jefferson@tecnun.com.br
' Data......: 20/08/2013
' Empresa...: Tecnun Tecnologia em Informática
' Descrição.: Rotina para gerar uma chave com o delimitador "|"
'---------------------------------------------------------------------------------------
Public Function GerarChaves(ParamArray Parametros() As Variant)
On Error GoTo TratarErro
Dim contador        As Integer
Dim ChaveAux        As String

    If UBound(Parametros) > 0 Then
        For contador = 0 To UBound(Parametros) Step 1
            ChaveAux = ChaveAux & Parametros(contador) & "|"
        Next contador
        ChaveAux = VBA.Left(ChaveAux, VBA.Len(ChaveAux) - 1)
    Else
        ChaveAux = Parametros(0)
    End If
    GerarChaves = ChaveAux
On Error GoTo 0
Exit Function
TratarErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Publicas.GerarChaves", Erl)
End Function

Public Function ExibirAvisos()
    Call AbrirFormulario("frmAvisos")
End Function


Function FormataCPFBNPJ(pValor As String) As String
    Dim pCodigo As String
10  pCodigo = pValor
20  pCodigo = VBA.Replace(pCodigo, "-", "")
30  pCodigo = VBA.Replace(pCodigo, "(", "")
40  pCodigo = VBA.Replace(pCodigo, ")", "")
50  pCodigo = VBA.Replace(pCodigo, " ", "")
60  Select Case Len(pCodigo)
    Case 11: pCodigo = VBA.Format(pCodigo, "000\.000\.000\-00")
70  Case 14: pCodigo = VBA.Format(pCodigo, "00\.000\.000/0000-00")
80  Case Else: pCodigo = pValor
90  End Select
100 FormataCPFBNPJ = pCodigo
End Function

Function FormataCEP(pValor As Variant) As String
    FormataCEP = VBA.Format(pValor, "00000\-000")
End Function

Function FormataCNPJ(pValor As Variant) As String
    FormataCNPJ = VBA.Format(pValor, "00\.000\.000/0000-00")
End Function

Function FormataTelefone(pNumTel As Variant, Optional bIncluirParenteses As Boolean) As String
    Dim abre_parentese As String, fecha_parentese As String
    pNumTel = VBA.Replace(pNumTel, "-", "")
    pNumTel = VBA.Replace(pNumTel, "(", "")
    pNumTel = VBA.Replace(pNumTel, ")", "")
    pNumTel = VBA.Replace(pNumTel, " ", "")
    
    If bIncluirParenteses Then abre_parentese = "("
    If bIncluirParenteses Then fecha_parentese = ")"
    
    Select Case Len(pNumTel)
    Case 8
        pNumTel = VBA.Format(pNumTel, "0000\-0000")
    Case 9
        pNumTel = VBA.Format(pNumTel, " 00000\-0000") 'Padrão 9 Digitos
    Case 11, 10 'Somente com o DDD
        If Mid(pNumTel, 3, 1) = "9" Then
            pNumTel = VBA.Format(pNumTel, abre_parentese & "00" & fecha_parentese & " 00000\-0000")
        Else
            pNumTel = VBA.Format(pNumTel, abre_parentese & "00" & fecha_parentese & " 0000\-0000") 'Padrão 9 Digitos
        End If
    Case 13, 12 'Com o DDD e Código da Operadora
        If Mid(pNumTel, 5, 1) = "9" Then
            pNumTel = VBA.Format(pNumTel, abre_parentese & "00 00" & fecha_parentese & " 00000\-0000")
        Else
            pNumTel = VBA.Format(pNumTel, abre_parentese & "00 00" & fecha_parentese & " 0000\-0000") 'Padrão 9 Digitos
        End If
    Case Else
        FormataTelefone = pNumTel
    End Select
    FormataTelefone = "+55 " & pNumTel
End Function

Public Function InicializaConexaoDB(ProjectConnection As Object) As ConexaoDB
    Dim cDB As ConexaoDB
    Set cDB = New ConexaoDB
    Set cDB.ConexaoBanco = ProjectConnection
    'Reinicializa a coleção de procedures modelo
    Set cDB.ProceduresModelo = Nothing
    Call cDB.CarregarProceduresModelo
    Set InicializaConexaoDB = cDB
End Function

Sub ExportarElementosVBE()
    Dim fso             As Object
    Dim vbProject       As Object
    Dim vbComponent     As Object
    Dim pasta           As String
    Dim destino         As String
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    Set vbProject = VBE.VBProjects(1)
    pasta = fso.GetParentFolderName(vbProject.FileName)
    pasta = pasta & "\VBComponents"
    If Not fso.FolderExists(pasta) Then fso.CreateFolder pasta
    
    For Each vbComponent In vbProject.VBComponents
        With vbComponent
            destino = pasta & "\" & .Name
            Select Case .Type
                Case 1
                    destino = destino & ".bas"
                Case 2
                    destino = destino & ".cls"
                Case 100
                    destino = destino & ".frm"
            End Select
            Debug.Print .Name, .Type
            .export destino
        End With
    Next vbComponent
    
    Set fso = Nothing
    Set vbProject = Nothing
    Set vbComponent = Nothing
    
End Sub

Function FullTimeFormat(tmValue, Optional bIncludeSecs As Boolean) As String
    Dim X
    X = Int(CSng(CDate(tmValue) * 24)) & ":" & VBA.Format(tmValue, "nn" & VBA.IIf(bIncludeSecs, ":ss", ""))
    FullTimeFormat = X
End Function


'---------------------------------------------------------------------------------------
' Modulo....: Publicas/
' Rotina....: ColetaDadosLstBox / Function
' Autor.....: Victor Félix
' Contato...: victor.santos@mondial.com.br
' Data......: 24/07/2013
' Empresa...: Mondial Tecnologia em Informática LTDA.
' Descrição.: Coleta os items de uma ListBox Generica e adiciona-os em um dicionario
'---------------------------------------------------------------------------------------
Public Function ColetaDadosLstBox(ByVal accControl As Access.ListBox) As Object 'Scripting.Dictionary
On Error GoTo TrataErro
Dim dicAux              As Object
Dim lngContador         As Long
    Set dicAux = VBA.CreateObject("Scripting.Dictionary")
    For lngContador = 0 To accControl.ListCount
        If Not dicAux.Exists(accControl.ItemData(lngContador)) Then
            Call dicAux.Add(accControl.ItemData(lngContador), Nothing)
        End If
    Next lngContador
    Set ColetaDadosLstBox = dicAux
Exit Function
TrataErro:
    Call Excecoes.TratarErro(VBA.Err.Description, VBA.Err.Number, "Publicas.ColetaDadosLstBox", 0, True)
End Function

Function GetDocument() As Object
    Dim oApp As Object
    Set oApp = Application
    If oApp.Name = "Microsoft Excel" Then
        Set GetDocument = oApp.ThisWorkbook
    ElseIf oApp.Name = "Microsoft Access" Then
        Set GetDocument = oApp.CurrentDb
    Else
        Set GetDocument = oApp.getHostPath
    End If
    Set oApp = Nothing
End Function
