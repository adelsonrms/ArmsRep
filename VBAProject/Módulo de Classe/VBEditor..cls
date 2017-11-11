'---------------------------------------------------------------------------------------
' MÓDULO           : Projeto.VBEditor
' TIPO             : Módulo de classe
' DATA/HORA        : 02/05/2008 15:30
' CONSULTOR        : Adelson Rosendo Marques da Silva
' DESCRIÇÃO        : Contem as funções necessárias para analise de codigo
'---------------------------------------------------------------------------------------
' + NOTAS
' Documentação
' Continuação da versão do código de 2008
' Aprimoramento da codificação, Limpeza e organização
' Analisa todos os projetos na inicailização
'
' + HISTÓRIO DE REVISÃO
'---------------------------------------------------------------------------------------
' DATA / DESCRIÇÃO
'---------------------------------------------------------------------------------------
' 02/05/2008 15:30 - Versão 1.0 Implementação inicial do codigo
' 12/09/2014 00:14 - Versão 2.0 do codigo
' 10/11/2017 22:14 - Versão 3.0 - Revisão e ajustes no código. Alteração dos nomes dos objetos para portugues
'---------------------------------------------------------------------------------------
Public Enum eTipoPropriedade
    Proc_Function = 0
    Property_Let = 1
    Property_set = 2
    Property_Get = 3
End Enum

Public Enum eTypeModifier
    mdf_Public = 0
    mdf_Private = 1
    mdf_Friend = 2
    mdf_Static = 3
    mdf_Dim = 4
    mdf_Default = 0
End Enum

Public Enum eComponentType
    CompType_Modulo = 1
    CompType_Classe = 2
    CompType_Form = 3
    CompType_Documento = 100
End Enum

Public Enum eVBDataType
    vbdt_Byte = 0
    vbdt_Boolean = 1
    vbdt_Integer = 2
    vbdt_Long = 3
    vbdt_Single = 4
    vbdt_Double = 5
    vbdt_Currency = 6
    vbdt_Decimal = 7
    vbdt_Date = 8
    vbdt_Objeto = 9
    vbdt_String = 10
    vbdt_Variant = 11
End Enum

Public Enum eProcedureType
    procType_Const = 0
    procType_Declare = 1
    procType_Enum = 2
    procType_Function = 3
    procType_Property = 4
    procType_Sub = 5
    procType_Type = 6
End Enum

Private m_ObjVBE As VBIDE.VBE
Private m_Objapp As Object
Private m_objProjetos As Projetos

Private m_sPastaBackup As String


'Determina o tipo de Aplicação
Private Sub DeterminiHost()
    On Error Resume Next
    Set VBEHost = Application.VBE
End Sub

'Inicializa a classe.
'Analisa o código de todos os projetos
Private Sub Class_Initialize()
    DeterminiHost
    Set m_objProjetos = CarregarProjetos()
End Sub

Public Property Get PastaBackup() As String
    PastaBackup = m_sPastaBackup
End Property

Public Property Let PastaBackup(ByVal sNewValue As String)
    m_sPastaBackup = sNewValue
End Property

Public Property Get Projetos() As Projetos
    Set Projetos = m_objProjetos
End Property

Private Sub Class_Terminate()
    Set m_objProjetos = Nothing
End Sub

Public Property Get app() As Object
    Set app = m_Objapp
End Property

Public Property Set app(ByVal ObjValue As Object)
    Set m_Objapp = ObjValue
End Property

Public Property Get VBEHost() As VBIDE.VBE
    Set VBEHost = m_ObjVBE
End Property

Public Property Set VBEHost(ByVal ObjValue As VBIDE.VBE)
    Set m_ObjVBE = ObjValue
End Property

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : VBEditor.GetDeclarationCode()
' TIPO             : Function
' DATA/HORA        : 12/09/2014 00:02
' CONSULTOR        : Adelson Rosendo Marques da Silva
' CONTATO          : adelsons@gmail.com.br
' DESCRIÇÃO        : Retorena o codigo de declaração de um módulo
'---------------------------------------------------------------------------------------
Private Function GetDeclarationCode(cmCodeModule As Variant) As String

    '## MZTools Custom - Variáveis de Ambiente
    '----------------------------------------------------------------------------------------------------
    On Error GoTo GetDeclarationCode_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
    Dim dtSartRunProc As Date: dtSartRunProc = Time
    Const cstr_ProcedureName As String = "Function Projeto.VBEditor.GetDeclarationCode()"

    On Error Resume Next
    Dim VBCodeMod As CodeModule
    'Determina o Variavel 'CodeModulo' a ser utilizado
    'Se for informado o Variavel, pega - o diretamente
    If TypeName(cmCodeModule) = "CodeModule" Then
        Set VBCodeMod = cmCodeModule
        'Se for informado uma String, pega o codemodule do VBComponent
    ElseIf VarType(cmCodeModule) = vbString Then
        Set VBCodeMod = VBEHost.ActiveVBProject.VBComponentes(cmCodeModule).CodeModule
    End If
    'Primeira linha após a ultima linha da seção de declaração
    GetDeclarationCode = VBCodeMod.Lines(1, VBCodeMod.CountOfDeclarationLines)
    
    On Error GoTo 0
    Exit Function

GetDeclarationCode_Error:
    If Err <> 0 Then
        Call MsgBox("An error ocurred at running code !" & vbNewLine & vbNewLine & _
            "Code Line : " & VBA.Erl & "\ Error > " & Err.Number & "-" & Err.Description, vbCritical, "ARMS VBE Code Analyzer 1.0 - Error")
    End If
    Exit Function
    'Debug Mode
    Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : VBEditor.GetComponents()
' TIPO             : Function
' DATA/HORA        : 11/09/2014 23:59
' CONSULTOR        : Adelson Rosendo Marques da Silva
' CONTATO          : adelsons@gmail.com.br
' DESCRIÇÃO        : Analisa todos os componentes de um projeto
'---------------------------------------------------------------------------------------
Private Function GetComponents(vbProj As VBProject)
    '----------------------------------------------------------------------------------------------------
    On Error GoTo TratarErro: Const cstr_ProcedureName As String = "VBAProject.VBEditor.GetComponents"
    '------------------------------------------------------------------------------------------------
    Dim vbac   As Componente
    Dim vbc    As VBComponent
    Dim vbacs  As New Componentes

    If vbProj.Protection = vbext_pp_none Then
        For Each vbc In vbProj.VBComponents
            Set vbac = New Componente
            With vbac
                .DeclarationCode = GetDeclarationCode(vbc.CodeModule)
                .nome = vbc.Name
                .Tipo = vbc.Type
                .TipoNome = GetComponentType_Name(vbc.Type)
                .Extensao = GetExtensionByType(vbc.Type)
                Set .Procedimentos = GetProcedures(vbc.CodeModule)
                Set .VariaveisGlobais = GetVariables(.DeclarationCode)
                Set .VBComponentParent = vbc
            End With
            Call vbacs.add(vbac, vbac.nome)
        Next
    End If
    
    Set GetComponents = vbacs
    
    On Error GoTo 0
Fim:
    Exit Function
TratarErro:
    If VBA.Err <> 0 Then Call VBA.Err.Raise(VBA.Rnd(), cstr_ProcedureName & " - " & VBA.Erl() & ">" & VBA.Err.Description & VBA.vbCrLf)
    GoTo Fim:
    Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : VBEditor.GetProcedures()
' TIPO             : Function
' DATA/HORA        : 12/09/2014 00:04
' CONSULTOR        : Adelson Rosendo Marques da Silva
' CONTATO          : adelsons@gmail.com.br
' DESCRIÇÃO        : Analisa um componente e retorna a lista de procedimentos
'---------------------------------------------------------------------------------------
Private Function GetProcedures(cmCodeModule As Variant) As Procedimentos

    '## MZTools Custom - Variáveis de Ambiente
    '----------------------------------------------------------------------------------------------------
    On Error GoTo GetProcedures_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
    Dim dtSartRunProc As Date: dtSartRunProc = Time
    Const cstr_ProcedureName As String = "Function Projeto.VBEditor.GetProcedures()"
    '----------------------------------------------------------------------------------------------------

    On Error Resume Next
    Dim QtdLinhasPorProcedimento As Long
    Dim VBCodeMod As CodeModule
    Dim intLinha As Long

    Dim strNomeProc As String
    Dim TipoProcedimento As eProcedureType
    Dim strModificador As String

    Dim strLinhaDeclaração As String
    'Instancia da classe
    Dim oProcedimento As Procedimento
    'Variável do tipo de procedimento
    Dim vbTipoProc As vbext_ProcKind
    '
    Dim TipoPropriedade As eTipoPropriedade
    Dim iTipoFunção As Integer

    Set GetProcedures = New Procedimentos

    'Determina o Variavel 'CodeModulo' a ser utilizado
    'Se for informado o Variavel, pega - o diretamente
    If TypeName(cmCodeModule) = "CodeModule" Then
        Set VBCodeMod = cmCodeModule
        'Se for informado uma String, pega o codemodule do VBComponent
    ElseIf VarType(cmCodeModule) = vbString Then
        Set VBCodeMod = VBEHost.ActiveVBProject.VBComponentes(cmCodeModule).CodeModule
    End If

    With VBCodeMod
        'Primeira linha após a ultima linha da seção de declaração
        intLinha = .CountOfDeclarationLines + 1
        'Executa até que a linha seja maior que o total de linhas de codigo do modulo
        Do Until intLinha >= .CountOfLines
            'Determina a linha de declaração
            strLinhaDeclaração = .Lines(intLinha, 1)
            TipoProcedimento = GetProcedureType(strLinhaDeclaração)
            If TipoProcedimento = 0 Then GoTo ProximaLinha
            'Determine o modificador
            strModificador = GetProcedureScopeModifier(strLinhaDeclaração)

            'Determina que tipo de procedimento será usada
            Select Case TipoProcedimento
                Case eProcedureType.procType_Property
                    vbTipoProc = GetPropertyType(strLinhaDeclaração)
                    TipoPropriedade = GetPropertyType(strLinhaDeclaração)
                Case Else
                    vbTipoProc = vbext_pk_Proc
            End Select

            'Retorna o nome do procedimento
            strNomeProc = VBA.Trim(.ProcOfLine(intLinha, vbTipoProc))
            'Retorna o total de linhas de codigo
            QtdLinhasPorProcedimento = VBCodeMod.ProcCountLines(strNomeProc, vbTipoProc)
            'Cria uma nova Instancia da classe Procedimento
            Set oProcedimento = New Procedimento
            'Preenche as informações na classe
            With oProcedimento
                .TipoPropriedade = vbTipoProc
                .Declaração = strLinhaDeclaração
                .nome = strNomeProc
                .Tipo = TipoProcedimento
                .TypeNamo = GetDataType_Name(.Tipo)
                If TipoProcedimento = procType_Function Or (eProcedureType.procType_Property And vbTipoProc = vbext_pk_Get) Then
                    .Retorna = GetReturnProcedure(strLinhaDeclaração)
                Else
                    .Retorna = "N/A"
                End If
                .Modificador = strModificador
                .Código = VBCodeMod.Lines(VBCodeMod.ProcStartLine(.nome, vbTipoProc), VBCodeMod.ProcCountLines(.nome, vbTipoProc))
                Set .Argumentos = GetArguments(strLinhaDeclaração, oProcedimento)

                iTipoFunção = TipoProcedimento

                Select Case TipoProcedimento
                    Case eProcedureType.procType_Property
                        .TipoMetodo = GetProcedureType_Name(iTipoFunção) & " " & GetPropertyType_Name(TipoPropriedade)
                    Case Else
                        .TipoMetodo = GetProcedureType_Name(iTipoFunção)
                End Select

                Set .Variaveis = GetVariables(.Código, oProcedimento)

                If Not GetProcedures.Exists(vbTipoProc & strNomeProc) Then GetProcedures.add oProcedimento, vbTipoProc & strNomeProc
            End With
ProximaLinha:
            'Procima linha de procedimento
            intLinha = intLinha + 1
        Loop
    End With

    Set GetProcedures = GetProcedures

    On Error GoTo 0
    Exit Function

GetProcedures_Error:
    If Err <> 0 Then
        Call MsgBox("An error ocurred at running code !" & vbNewLine & vbNewLine & _
            "Code Line : " & VBA.Erl & "\ Error > " & Err.Number & "-" & Err.Description, vbCritical, "ARMS VBE Code Analyzer 1.0 - Error")
    End If
    Exit Function
    'Debug Mode
    Resume

End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : VBEditor.GetVariables()
' TIPO             : Function
' DATA/HORA        : 17/1/2009 - 03:02
' CONSULTOR        : Adelson Rosendo Marques da Silva
' CONTATO          : adelsons@gmail.com.br
' DESCRIÇÃO        : Analisa o código de um componente e recupera suas variáveis
'---------------------------------------------------------------------------------------
Private Function GetVariables(strVBCode As String, Optional parent As Procedimento) As Variaveis
    Dim LinhaCodigo As String
    Dim CodigoVB As Variant
    Dim intLinha As Long
    Dim oVariavel As Variavel
    Dim varNome As String
    Dim vatTipo As String
    
    '## MZTools Custom - Variáveis de Ambiente
    '----------------------------------------------------------------------------------------------------
    On Error GoTo GetVariables_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
    Dim dtSartRunProc As Date: dtSartRunProc = Time
    Const cstr_ProcedureName As String = "Function Projeto.VBEditor.GetVariables()"
    '----------------------------------------------------------------------------------------------------

    Set GetVariables = New Variaveis
    'Divide o codigo em linhas
    CodigoVB = VBA.Split(strVBCode, vbNewLine)

    'Percurre todas as linhas de codigo
    For intLinha = LBound(CodigoVB) To UBound(CodigoVB)
        'Cria uma nova instancia da classe Variável
        Set oVariavel = New Variavel
        'Pega a linha Atual
        LinhaCodigo = VBA.Trim(CodigoVB(intLinha))    'DeclaraçãoMultipla(intDec)
        If Not IsDeclarationLine(LinhaCodigo) Then
            If LinhaCodigo <> "" Then
                Set oVariavel = GetVarInfo(LinhaCodigo)
                If oVariavel.nome <> "" Then
                    Set oVariavel.parent = parent
                    oVariavel.Modifier = GetProcedureScopeModifier_Name(GetVariableModifier(LinhaCodigo))
                    If Not GetVariables.Exists(oVariavel.nome) Then
                        GetVariables.add oVariavel, oVariavel.nome
                    End If
                End If
            End If
        End If
    Next

    Set GetVariables = GetVariables

    On Error GoTo 0
    Exit Function

GetVariables_Error:
    If Err <> 0 Then
        Call MsgBox("An error ocurred at running code !" & vbNewLine & vbNewLine & _
            "Code Line : " & VBA.Erl & "\ Error > " & Err.Number & "-" & Err.Description, vbCritical, "ARMS VBE Code Analyzer 1.0 - Error")
    End If
    Exit Function
    'Debug Mode
    Resume
End Function

'Verifica se a linha de codigo é uma linha de declaração
Private Function IsDeclarationLine(strLinhaDeclaração As String) As Boolean
    Dim PalavraChave As String
    Dim PalavraChave2 As String
    vPrimeraPalavra = VBA.Split(strLinhaDeclaração, " ")
    If UBound(vPrimeraPalavra) < 1 Then Exit Function
    'Retorna a primeira palavra
    PalavraChave = VBA.Trim(vPrimeraPalavra(0))
    IsDeclarationLine = IsProcedure(PalavraChave)
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : VBEditor.GetVarInfo()
' TIPO             : Function
' DATA/HORA        : 19/1/2009 - 17:09
' CONSULTOR        : Adelson Rosendo Marques da Silva
' CONTATO          : adelsons@gmail.com.br
' DESCRIÇÃO        : Recupera informações sobre uma variável
'---------------------------------------------------------------------------------------
Private Function GetVarInfo(LinhaCodigo As String) As Variavel
    Dim lngDim As Long
    Dim lngAs  As Long
    Dim NomeETipo As Variant
    Dim ChecaNome
    Dim DeclaraçãoMultipla
    Dim intDec As Integer

    '## MZTools Custom - Variáveis de Ambiente
    '----------------------------------------------------------------------------------------------------
    On Error GoTo GetVarInfo_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
    Dim dtSartRunProc As Date: dtSartRunProc = Time
    Const cstr_ProcedureName As String = "Function Projeto.VBEditor.GetVarInfo()"
    '----------------------------------------------------------------------------------------------------

    Set GetVarInfo = New Variavel
    LinhaCodigo = VBA.Trim(LinhaCodigo)
    'Procura pela instrução Dim
    lngDim = InStr(1, LinhaCodigo, "Dim ")

    If lngDim = 0 Then
        lngDim = InStr(1, LinhaCodigo, "Private ")
    End If

    If lngDim = 0 Then
        lngDim = InStr(1, LinhaCodigo, "Public ")
    End If

    'Procura pela clausula
    lngAs = InStr(1, LinhaCodigo, " As ")

    If lngAs > 0 Then
        NomeETipo = VBA.Split(LinhaCodigo, " As ")
        ChecaNome = VBA.Split(VBA.Trim(NomeETipo(LBound(NomeETipo))), " ")

        If UBound(ChecaNome) <> 1 Then Exit Function
        'Retorna o nome
        GetVarInfo.nome = VBA.Trim(ChecaNome(1))
        'Procura pela palavra New
        vNew = VBA.Split(NomeETipo(UBound(NomeETipo)), "New ")
        'Se encontrado , é um Objeto
        If UBound(vNew) > 0 Then
            'Defini como true
            GetVarInfo.IsObject = True
            'Pega o nome do objeto
            GetVarInfo.ObjectName = VBA.Trim(VBA.Split(NomeETipo(UBound(NomeETipo)), " ")(1))
            GetVarInfo.TipoDeDados = "Object"
        Else

            GetVarInfo.TipoDeDados = VBA.Trim(NomeETipo(UBound(NomeETipo)))
            If Not GetVarInfo.TipoDeDados = GetDataType_Name(GetDataTypeIndex(GetVarInfo.TipoDeDados)) Then
                GetVarInfo.IsObject = True
                GetVarInfo.ObjectName = GetVarInfo.TipoDeDados
                GetVarInfo.TipoDeDados = "Object"
            Else
                GetVarInfo.ObjectName = "(Não é objeto)"
                GetVarInfo.IsObject = False
            End If

        End If
    ElseIf lngDim > 0 Then
        GetVarInfo.nome = VBA.Trim(VBA.Split((LinhaCodigo), " ")(1))
        GetVarInfo.IsObject = True
        GetVarInfo.ObjectName = GetVarInfo.TipoDeDados
        GetVarInfo.TipoDeDados = "Object"
    End If

    On Error GoTo 0
    Exit Function

GetVarInfo_Error:
    If Err <> 0 Then
        Call MsgBox("An error ocurred at running code !" & vbNewLine & vbNewLine & _
            "Code Line : " & VBA.Erl & "\ Error > " & Err.Number & "-" & Err.Description, vbCritical, "ARMS VBE Code Analyzer 1.0 - Error")
    End If
    Exit Function
    'Debug Mode
    Resume
End Function

Private Function GetEscopeVar(strType As String)
    Select Case strType
        Case "Dim"
            GetEscopeVar = "Local"
    End Select
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : VBEditor.GetArguments()
' TIPO             : Function
' DATA/HORA        : 12/09/2014 00:10
' CONSULTOR        : Adelson Rosendo Marques da Silva
' CONTATO          : adelsons@gmail.com.br
' DESCRIÇÃO        : Recupera a lista de argumentos de uma função/sub
'---------------------------------------------------------------------------------------
Private Function GetArguments(strLinhaDeclaração, Optional ByVal parent As Procedimento) As Argumentos
    Dim lngAbre_Parent As Long
    Dim lngFecha_Parent As Long
    Dim strDefArgumentos As String
    Dim vArgs  As Variant
    Dim vDetalhes As Variant

    Dim oArgumento As Argumento


    '## MZTools Custom - Variáveis de Ambiente
    '----------------------------------------------------------------------------------------------------
    On Error GoTo GetArguments_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
    Dim dtSartRunProc As Date: dtSartRunProc = Time
    Const cstr_ProcedureName As String = "Function Projeto.VBEditor.GetArguments()"
    '----------------------------------------------------------------------------------------------------

    Set GetArguments = New Argumentos

    lngAbre_Parent = InStr(1, strLinhaDeclaração, "(")
    lngFecha_Parent = InStr(lngAbre_Parent, strLinhaDeclaração, ")")
    If (lngFecha_Parent - lngAbre_Parent) = 1 Then Exit Function
    strDefArgumentos = Mid$(strLinhaDeclaração, lngAbre_Parent + 1, (lngFecha_Parent - lngAbre_Parent) - 1)
    vArgs = VBA.Split(strDefArgumentos, ",")

    For i = 0 To UBound(vArgs)
        Set oArgumento = New Argumento
        vDetalhes = VBA.Split(VBA.Trim(vArgs(i)), " ")
        With oArgumento
            Set .parent = parent
            .IsOptional = vDetalhes(0) = "Optional"
            If InStr(vArgs(i), "=") > 0 Then
                .DefaulValue = VBA.Split(vArgs(i), " = ")(1)
            End If
            If InStr(vArgs(i), "ByVal") > 0 Then .ReferenceType = "ByVal"
            If InStr(vArgs(i), "ByRef") > 0 Then .ReferenceType = "ByRef"
            If .ReferenceType = "" Then .ReferenceType = "(Default)"

            If InStr(vArgs(i), " As ") > 0 Then
                vNome = VBA.Split(vArgs(i), " As ")
                If InStr(VBA.Trim(vNome(0)), " ") > 0 Then
                    .Name = VBA.Trim(Mid(VBA.Trim(vNome(0)), InStr(VBA.Trim(vNome(0)), " ") + 1))
                Else
                    .Name = vNome(0)
                End If

                If InStr(VBA.Trim(vNome(1)), " ") > 0 Then
                    .DataType = VBA.Trim(Left(vNome(1), InStr(vNome(1), " ")))
                Else
                    .DataType = vNome(1)
                End If
            Else
                .Name = VBA.Split(vArgs(i), " ")(UBound(VBA.Split(vArgs(i), " ")))
            End If
            If .DataType = "" Then .DataType = "Variant"
            GetArguments.add oArgumento, .Name
        End With
    Next

    On Error GoTo 0
    Exit Function

GetArguments_Error:
    If Err <> 0 Then
        Call MsgBox("An error ocurred at running code !" & vbNewLine & vbNewLine & _
            "Code Line : " & VBA.Erl & "\ Error > " & Err.Number & "-" & Err.Description, vbCritical, "ARMS VBE Code Analyzer 1.0 - Error")
    End If
    Exit Function
    'Debug Mode
    Resume
End Function
'--------------------------------------------------------------------------------------------------------------
' FUNCOES DE APOIO DE ANALISE DO CÓDIGO
'--------------------------------------------------------------------------------------------------------------
'Determina o retorno da função
Private Function GetReturnProcedure(strLinhaDeclaração As String)
    Dim lngAbre_Parent As Long
    Dim lngFecha_Parent As Long
    Dim strDefArgumentos As String
    Dim vDec As Variant
    Dim vDetalhes As Variant
    lngFecha_Parent = InStr(1, strLinhaDeclaração, ")") + 1
    vDec = VBA.Split(Mid(strLinhaDeclaração, lngFecha_Parent, Len(strLinhaDeclaração) - lngFecha_Parent + 1), "As")
    If UBound(vDec) <= 0 Then
        GetReturnProcedure = "Variant (Default)"
    Else
        GetReturnProcedure = VBA.Trim(vDec(UBound(vDec)))
    End If
End Function

Private Function GetPropertyType(strLinhaDeclaração As String) As eTipoPropriedade
    Dim PalavraChave As String
    vPrimeraPalavra = VBA.Split(strLinhaDeclaração, " ")
    If UBound(vPrimeraPalavra) < 1 Then Exit Function
    'Retorna a primeira palavra
    PalavraChave = VBA.Trim(vPrimeraPalavra(0))
    'Veriica se a palavra e um modificador
    If IsScopeModifier(PalavraChave) Then
        'Se sim, pega a segunda palavra que deve ser o tipo do procedimento
        PalavraChave = VBA.Trim(vPrimeraPalavra(2))
    End If
    Select Case VBA.Trim(VBA.LCase(PalavraChave))
        Case VBA.LCase("Get")
            GetPropertyType = Property_Get
        Case VBA.LCase("Let")
            GetPropertyType = Property_Let
        Case VBA.LCase("Set")
            GetPropertyType = Property_set
    End Select
End Function

Private Function GetProcedureScopeModifier_Name(vIndex As Integer) As String
    Dim ArrModifier()
    ArrModifier = Array("Public", "Private", "Friend", "Static", "Dim")
    GetProcedureScopeModifier_Name = ArrModifier(vIndex)
End Function

Private Function GetProcedureType_Name(vIndex As Integer) As String
    Dim ArrProcedureType()
    ArrType = Array("Const", _
        "Declare", _
        "Enum", _
        "Function", _
        "Property", _
        "Sub", _
        "Type")
    GetProcedureType_Name = ArrType(vIndex)
End Function
'Retorna o tipo de Propriedade
Private Function GetPropertyType_Name(vIndex As eTipoPropriedade) As String
    Dim ArrType()
    ArrType = Array("", "Let", "Set", "Get")
    GetPropertyType_Name = ArrType(vIndex)
End Function
'Retorna o indice do tipo de dados baseado no nome
Private Function GetDataTypeIndex(strDataType As String) As eVBDataType
    Select Case VBA.Trim(VBA.LCase(strDataType))
        Case VBA.LCase("Byte")
            GetDataTypeIndex = vbdt_Byte
        Case VBA.LCase("Boolean")
            GetDataTypeIndex = vbdt_Boolean
        Case VBA.LCase("Integer")
            GetDataTypeIndex = vbdt_Integer
        Case VBA.LCase("Long")
            GetDataTypeIndex = vbdt_Long
        Case VBA.LCase("Single")
            GetDataTypeIndex = vbdt_Single
        Case VBA.LCase("Double")
            GetDataTypeIndex = vbdt_Double
        Case VBA.LCase("Currency")
            GetDataTypeIndex = vbdt_Currency
        Case VBA.LCase("Decimal")
            GetDataTypeIndex = vbdt_Decimal
        Case VBA.LCase("Date")
            GetDataTypeIndex = vbdt_Date
        Case VBA.LCase("Objeto")
            GetDataTypeIndex = vbdt_Objeto
        Case VBA.LCase("String")
            GetDataTypeIndex = vbdt_String
        Case VBA.LCase("Variant")
            GetDataTypeIndex = vbdt_Variant
    End Select
End Function
'Retorna o nome do tipo de dado
Private Function GetDataType_Name(vIndex As eVBDataType) As String
    Dim ArrType()
    ArrType = Array("Byte", "Boolean", "Integer", "Long", "Single", "Double", "Currency", "Decimal", "Date", "Objeto", "String", "Variant")
    GetDataType_Name = ArrType(vIndex)
End Function
'Retorna o nome do tipo de componente baseado no indice
Private Function GetComponentType_Name(vIndex As Integer) As String
    Dim vbcom As vbext_ComponentType
    Dim ArrType()
    If vIndex = 11 Then vIndex = 4
    If vIndex = 100 Then vIndex = 5
    ArrType = VBA.Array("", "Módulo Padrão", "Módulo de Classe", "MSForm", "ActiveXDesigner", "Documento Excel")
    GetComponentType_Name = ArrType(vIndex)
End Function


Private Function GetExtensionByType(vIndex As Integer) As String
    Dim ArrType()
    If vIndex = 11 Then vIndex = 4
    If vIndex = 100 Then vIndex = 5
    ArrType = VBA.Array("", ".bas", ".cls", ".frm", ".axc", ".txt")
    GetExtensionByType = ArrType(vIndex)
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : GetVariableModifier
' TIPO             : Function
' DATA/HORA        : 24/12/2008 - 01:48
' DESENVOLVEDOR    : Adelson Rosendo Marques da Silva
' DESCRIÇÃO        : Retorna o tipo de modificador do procedimento
'---------------------------------------------------------------------------------------
Private Function GetVariableModifier(strLinhaDeclaração As String) As eTypeModifier
    '----------------------------------------------------------------------------------------------------
    On Error GoTo TratarErro: Const cstr_ProcedureName As String = "VBAProject.VBEditor.GetVariableModifier"
    '------------------------------------------------------------------------------------------------
    Dim PalavraChave As String
    If strLinhaDeclaração = "" Then Exit Function
    vPrimeraPalavra = VBA.Split(strLinhaDeclaração, " ")

    'Retorna a primeira palavra
    PalavraChave = VBA.Trim(vPrimeraPalavra(0))
    'Veriica se a palavra e um tipo de procedimento, por exemplo 'Sub' ou 'Function'
    If IsProcedure(PalavraChave) Then
        GetVariableModifier = mdf_Default
        Exit Function
    End If

    Select Case VBA.Trim(VBA.LCase(vPrimeraPalavra(0)))
        Case VBA.LCase("Public")
            GetVariableModifier = mdf_Public
        Case VBA.LCase("Private")
            GetVariableModifier = mdf_Private
        Case VBA.LCase("Friend")
            GetVariableModifier = mdf_Friend
        Case VBA.LCase("Static")
            GetVariableModifier = mdf_Static
        Case VBA.LCase("Dim")
            GetVariableModifier = mdf_Dim
    End Select
    
    On Error GoTo 0
Fim:
    Exit Function
TratarErro:
    If VBA.Err <> 0 Then Call VBA.Err.Raise(VBA.Rnd(), cstr_ProcedureName & " - " & VBA.Erl() & ">" & VBA.Err.Description & VBA.vbCrLf)
    GoTo Fim:
    Resume
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : GetProcedureScopeModifier
' TIPO             : Function
' DATA/HORA        : 24/12/2008 - 01:48
' DESENVOLVEDOR    : Adelson Rosendo Marques da Silva
' DESCRIÇÃO        : Retorna o tipo de modificador do procedimento
'---------------------------------------------------------------------------------------
Private Function GetProcedureScopeModifier(strLinhaDeclaração As String) As eTypeModifier
    '----------------------------------------------------------------------------------------------------
    On Error GoTo TratarErro: Const cstr_ProcedureName As String = "VBAProject.VBEditor.GetProcedureScopeModifier"
    '------------------------------------------------------------------------------------------------
    'Separa a linha de código por espaço em branco
    'Trata a ocorrência de erros na Procedimento
    Dim PalavraChave As String
    vPrimeraPalavra = VBA.Split(strLinhaDeclaração, " ")

    'Retorna a primeira palavra
    PalavraChave = VBA.Trim(vPrimeraPalavra(0))
    'Veriica se a palavra e um tipo de procedimento, por exemplo 'Sub' ou 'Function'
    If IsProcedure(PalavraChave) Then
        GetProcedureScopeModifier = mdf_Default
        Exit Function
    End If

    Select Case VBA.Trim(VBA.LCase(vPrimeraPalavra(0)))
        Case VBA.LCase("Public")
            GetProcedureScopeModifier = mdf_Public
        Case VBA.LCase("Private")
            GetProcedureScopeModifier = mdf_Private
        Case VBA.LCase("Friend")
            GetProcedureScopeModifier = mdf_Friend
        Case VBA.LCase("Static")
            GetProcedureScopeModifier = mdf_Static
        Case VBA.LCase("Dim")
            GetProcedureScopeModifier = mdf_Dim
    End Select
    
    On Error GoTo 0
Fim:
    Exit Function
TratarErro:
    If VBA.Err <> 0 Then Call VBA.Err.Raise(VBA.Rnd(), cstr_ProcedureName & " - " & VBA.Erl() & ">" & VBA.Err.Description & VBA.vbCrLf)
    GoTo Fim:
    Resume
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : GetProcedureType
' TIPO             : Function
' DATA/HORA        : 24/12/2008 - 01:48
' DESENVOLVEDOR    : Adelson Rosendo Marques da Silva
' DESCRIÇÃO        : Retorna o tipo de procedimento ou estrutura de declaração de variável
'---------------------------------------------------------------------------------------
Private Function GetProcedureType(strLinhaDeclaração As String) As eProcedureType
    'Separa a linha de código por espaço em branco
    'Trata a ocorrência de erros na Procedimento
    Dim PalavraChave As String

    '## MZTools Custom - Variáveis de Ambiente
    '----------------------------------------------------------------------------------------------------
    On Error GoTo GetProcedureType_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
    Dim dtSartRunProc As Date: dtSartRunProc = Time
    Const cstr_ProcedureName As String = "Function Projeto.VBEditor.GetProcedureType()"
    '----------------------------------------------------------------------------------------------------
   
    vPrimeraPalavra = VBA.Split(strLinhaDeclaração, " ")

    If UBound(vPrimeraPalavra) < 1 Then Exit Function

    'Retorna a primeira palavra
    PalavraChave = VBA.Trim(vPrimeraPalavra(0))
    'Veriica se a palavra e um modificador
    If IsScopeModifier(PalavraChave) Then
        'Se sim, pega a segunda palavra que deve ser o tipo do procedimento
        PalavraChave = VBA.Trim(vPrimeraPalavra(1))
    End If
    'Avaliza o tipo retornado pela variável
    'Devolve o tipo de procedimento à função
    Select Case VBA.LCase(PalavraChave)
        Case VBA.LCase("Const")
            GetProcedureType = procType_Const
        Case VBA.LCase("Declare")
            GetProcedureType = procType_Declare
        Case VBA.LCase("Enum")
            GetProcedureType = procType_Enum
        Case VBA.LCase("Function")
            GetProcedureType = procType_Function
        Case VBA.LCase("Property")
            GetProcedureType = procType_Property
        Case VBA.LCase("Sub")
            GetProcedureType = procType_Sub
        Case VBA.LCase("Type")
            GetProcedureType = procType_Type
    End Select

    On Error GoTo 0
    Exit Function

GetProcedureType_Error:
    If Err <> 0 Then
        Call MsgBox("An error ocurred at running code !" & vbNewLine & vbNewLine & _
            "Code Line : " & VBA.Erl & "\ Error > " & Err.Number & "-" & Err.Description, vbCritical, "ARMS VBE Code Analyzer 1.0 - Error")
    End If
    Exit Function
    'Debug Mode
    Resume
End Function
'-----------------------------------------------------------------------------------------------
' PROCEDIMENTO     : IsScopeModifier
' TIPO             : Function
' DATA/HORA        : 24/12/2008 - 01:49
' DESENVOLVEDOR    : Adelson Rosendo Marques da Silva
' DESCRIÇÃO        : Verifica se a palavra chave informada é um tipo de modificador
'                    Em VB os modificadores válidos são: 'Private', Public' 'Friend' e 'Static'
'-----------------------------------------------------------------------------------------------
Private Function IsScopeModifier(strKeyWord As String) As Boolean
    Select Case VBA.LCase(strKeyWord)
        Case VBA.LCase("Public"), VBA.LCase("Private"), VBA.LCase("Friend"), VBA.LCase("Static")
            IsScopeModifier = True
        Case Else
            IsScopeModifier = False
    End Select
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : IsProcedure
' TIPO             : Function
' DATA/HORA        : 24/12/2008 - 01:49
' DESENVOLVEDOR    : Adelson Rosendo Marques da Silva
' DESCRIÇÃO        : Verifica se a palavra chave informada é um nome de um tipo
'                    de procedimento que pode ser 'Sub', 'Function' ou 'Property'
'---------------------------------------------------------------------------------------
Private Function IsProcedure(strKeyWord As String) As Boolean
    Select Case VBA.LCase(strKeyWord)
        Case VBA.LCase("Sub"), VBA.LCase("Function"), VBA.LCase("Property")
            IsProcedure = True
        Case Else
            IsProcedure = False
    End Select
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : VBEditor.CarregarProjetos()
' TIPO             : Sub
' DATA/HORA        : 11/09/2014 23:53
' CONSULTOR        : Adelson Rosendo Marques da Silva
' CONTATO          : adelsons@gmail.com.br
' DESCRIÇÃO        : Atualiza a lista de Projetos
'---------------------------------------------------------------------------------------
Public Function CarregarProjetos()
    '----------------------------------------------------------------------------------------------------
    On Error GoTo TratarErro: Const cstr_ProcedureName As String = "VBAProject.VBEditor.CarregarProjetos"
    '------------------------------------------------------------------------------------------------
    Dim vbProj As VBProject
    Dim Projeto As Projeto
    Dim c_Projeto As New Projetos

    For Each vbProj In Application.VBE.VBProjects
        If (VBA.InStr(vbProj.BuildFileName, "\") > 0) Or (VBA.InStr(vbProj.BuildFileName, "/") > 0) Then
            Set Projeto = New Projeto
            With Projeto
                .Name = vbProj.Name
                Set .Componentes() = GetComponents(vbProj)
                Set .parent = Me
                Set .Projeto = vbProj
                .ProtectionMode = "Unlocked"
                If vbProj.Protection = vbext_pp_locked Then .ProtectionMode = "Project is Locked"
            End With
            c_Projeto.add Projeto, vbProj.Filename
        End If
    Next vbProj
    Set CarregarProjetos = c_Projeto
    
    On Error GoTo 0
Fim:
    Exit Function
TratarErro:
    If VBA.Err <> 0 Then Call VBA.Err.Raise(VBA.Rnd(), cstr_ProcedureName & " - " & VBA.Erl() & ">" & VBA.Err.Description & VBA.vbCrLf)
    GoTo Fim:
    Resume
End Function

Function ExportCode(pCode As String, toFileName As String)
    '----------------------------------------------------------------------------------------------------
    On Error GoTo TratarErro: Const cstr_ProcedureName As String = "Projeto.VBEditor.ExportCode"
    '------------------------------------------------------------------------------------------------
    If toFileName = "" Then toFileName = VBA.Environ("TEMP") & "\TEMP_TEXT_OUT.TXT"
    Open toFileName For Output As #1
    Print #1, pCode
    Close #1
    On Error GoTo 0
Fim:
    Exit Function
TratarErro:
    If VBA.Err <> 0 Then Call VBA.Err.Raise(VBA.Rnd(), cstr_ProcedureName & " - " & VBA.Erl() & ">" & VBA.Err.Description & VBA.vbCrLf)
    GoTo Fim:
    Resume
End Function

'-------------------------------------------------------------------------------------------------------------------------------
' PROCEDIMENTO     : Function BackCode / Retorna : Variant()
' DATA/HORA        : 10/11/2017 21:29
' CONSULTOR        : Adelson
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Realiza o backup de todo o código projeto
'-------------------------------------------------------------------------------------------------------------------------------
Public Function BackCode() As Variant()
    Dim Result() As Variant
    Dim strPasta As String
    Dim vbc As VBCodeConvert
    Dim Projeto As Projeto
    Dim Comp As Componente
    Dim Util As New cUtil
    
    Set Projeto = Me.Projetos.Item(ThisWorkbook.VBProject.Filename)
    For Each Comp In Projeto.Componentes
        If Comp.Tipo = vbext_ct_StdModule Or Comp.Tipo = vbext_ct_ClassModule Then
            strPastaBackup = Me.PastaBackup & "\" & Projeto.Name & "\" & Comp.TipoNome
            Call Util.CriarDiretorio(strPastaBackup)
            Call Me.ExportCode(Comp.Codigo, strPastaBackup & "\" & Comp.nome & "." & Comp.Extensao)
        End If
    Next Comp
End Function

