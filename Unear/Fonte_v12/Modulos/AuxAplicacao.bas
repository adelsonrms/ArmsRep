Attribute VB_Name = "AuxAplicacao"
Option Explicit
'---------------------------------------------------------------------------------------
' MÓDULO           : TFWCliente.AuxAplicacao
' TIPO             : Módulo
' DATA/HORA        : 09/06/2015 23:29
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Contem funções relativas ao aplicativo
'---------------------------------------------------------------------------------------
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora       Autor     Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      09/06/2015 23:29 23:29 - Adelson   Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Public Const cstr_TFW As String = "TFW.accdb"
Public TFWApp                               As cTFW_Aplicacao
Public Web                                  As New cTFW_WebUtil
Public oExcel                               As cTFW_Excel
Public DateUtil                             As New cTFW_DateUtil

Public dtFirstTime As Date
Public dtLastTime As Date
Public tempoDecorido As String
Public UserIDSession As String
Public appLog As New cTFW_Log
Private m_sChaveUsuario As String

Public Enum eEscopoVariavel
    varglobal = 1
    varUsuario = 2
End Enum

Public Enum eVersao
    Desenvolvimento = 0
    Homogolacao = 1
    Producao = 2
End Enum



Public Function PegarLocalTFWInstalado():
    On Error Resume Next
    PegarLocalTFWInstalado = Access.References!tfw.FullPath
    If VBA.Err <> 0 Then
        PegarLocalTFWInstalado = CurrentDb.Name
    End If
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxAplicacao.InicializarAplicação()
' TIPO             : Function
' DATA/HORA        : 09/06/2015 14:00
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Inicializa uma nova Instamcia da classe que retem informações sobre a conexão
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      09/06/2015 14:00    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Public Function InicializarAplicação()
    Dim rsApp As Object
    Dim vRetValidacao As Variant
1   On Error GoTo InicializarAplicação_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
2   Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function AuxAplicacao.InicializarAplicação()"
    '----------------------------------------------------------------------------------------------------
3   Call appLog.RegistrarLog(cstr_ProcedureName & "() - ## Inicializando rotina.." & VBA.String(20, "."))
    '-----------------------------------------------------------------------------------------------------
    'Inicializa o Splash
    '-----------------------------------------------------------------------------------------------------
29  Call appLog.RegistrarLog(VBA.Space(3) & "Inicializa a tela Splah")
30  Call Access.Application.DoCmd.Close(acForm, "frmIniciar", Access.acSaveNo)
31  Call Access.Application.DoCmd.openForm("frmIniciar", Access.acNormal)
32  Call salvaVariavelAplicacao("appVersion", NomeVersaoAtual)
33  Call salvaVariavelAplicacao("img_bg_inicio", PegaEnderecoConfiguracoes() & "\bg.png")
34  Call salvaVariavelAplicacao("logo_file_path", PegaEnderecoConfiguracoes() & "\logo.jpg")

    '-------------------------------------------------------------------------------------------------------
    'Inicializa informações da aplicação
    '-------------------------------------------------------------------------------------------------------
    ''### REMOVIDO TFW ### 320 Set ConexaoCliente = Nothing
35  Set rsApp = Conexao.PegarRS("Pegar_InformacoesAplicacao")
36  If Not rsApp Is Nothing Then
        'Preenche as informações da aplicação com base na tabela tblApp
37      Set TFWApp = New cTFW_Aplicacao
38      If Not rsApp.EOF Then
39          With TFWApp
40              .idAplicacao = rsApp!idAplicacao.value
41              .NomeAplicacao = rsApp!appName.value
42              .VersaoAplicacao = rsApp!appVersion.value
43              .Cliente = rsApp!CustomerName.value
44              .ArquivoIcone = rsApp!icon_file_path.value
45              .ArquivoLogo = rsApp!logo_file_path.value
46              .CorTarja = rsApp!color_bar.value
47              .NomeRibbonPrincipal = rsApp!TabPrincial.value
48          End With
49      End If
50      Call RemoverObjetosMemoria(rsApp)
51  Else
52      VBA.MsgBox "Não foi possivel carregar informações sobre a aplicação !", VBA.vbCritical, "Erro"
53  End If

54  Call ConfigurarAplicativo

    '-------------------------------------------------------------------------------------------------------
    'Inicializa a Ribbon
    '-------------------------------------------------------------------------------------------------------
55  Call InitializeRibbon

Fim:
56  Call appLog.RegistrarLog(cstr_ProcedureName & "() - ## Finalizado rotina / Tempo Total : " & appLog.CalcularTempoDecorrido(VBA.Now - dtSartRunProc) & String(20, "."))

57  On Error GoTo 0
58  Exit Function

InicializarAplicação_Error:
59  If VBA.Err <> 0 Then
60      If VBA.Err <> 32609 Then
61          lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
62          Call TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
63      Else
64          Resume Next
65      End If
66  End If
    GoTo Fim:
    'Debug Mode
67  Resume
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxAplicacao.ConfigurarAplicativo()
' TIPO             : Sub
' DATA/HORA        : 05/03/2015 10:52
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Configura opções padrões de um aplicativo
'                    dbText = 10
'                    dbBoolean = 1
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      05/03/2015 10:52    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Public Sub ConfigurarAplicativo()
    Carregar_Versao
    Call DefinirPropriedadeObjeto(CurrentDb, "AppTitle", pegarVariavelAplicacao("AppName") & " / " & pegarVariavelAplicacao("AppVersion"), 10, False)
80  Call DefinirPropriedadeObjeto(CurrentDb, "UseAppIconForFrmRpt", -1, 1, False)
    'Call DefinirPropriedadeObjeto(CurrentDb, "CustomRibbonId", "RibbonMain", dbText, False)
    Call DefinirPropriedadeObjeto(CurrentDb, "StartupShowStatusBar", -1, 1, False)
    Call ConfiguraIconeAplicativo
    Call Access.Application.RefreshTitleBar
End Sub
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxGlobal.ConfiguraIconeAplicativo()
' TIPO             : Sub
' DATA/HORA        : 03/09/2014 11:37
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Configura o icone do aplicativo
'---------------------------------------------------------------------------------------
Public Sub ConfiguraIconeAplicativo()
    Exit Sub
    
    
    'DESATIVADO. PROBLEMAS IDENTIFICADOS EM ALGUMAS MAQUINAS
    
    Dim strEndereco As String
    Dim dirPath As String
10  On Error GoTo ConfiguraIconeAplicativo_Error
20  dirPath = PegaEnderecoConfiguracoes
30  strEndereco = dirPath & "\icone.ico"
40  If Not AuxFileSystem.FileExists(strEndereco) Then
50      strEndereco = AuxDataBase.ExtrairAnexo("icon_app", "tblAnexos", dirPath, "icone.ico")
60  End If
70  strEndereco = AuxFileSystem.FormatPath(strEndereco)
90  Call DefinirPropriedadeObjeto(CurrentDb, "AppIcon", strEndereco, 10, False)

100 On Error GoTo 0
110 Exit Sub

ConfiguraIconeAplicativo_Error:
120 If VBA.Err <> 0 Then
130     Call RegistraLog("AuxGlobal.ConfiguraIconeAplicativo() - Linha : " & VBA.Erl & " / ERROR : " & VBA.Err.Number & "-" & VBA.Err.Description, CurrentProject.Path & "\Log\VBAErros.txt")
140 End If
150 Exit Sub
160 Resume
End Sub
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxGlobal.ListVBProjectReferences()
' TIPO             : Sub
' DATA/HORA        : 04/03/2015 12:15
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Lista/Enumera as referencias instaladas no projeto e guarda-as na tabela tblVBAReferences
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      04/03/2015 12:15    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Public Sub ListVBProjectReferences()
    Dim rsRefs As Object
    Dim Ref As Object

10  On Error GoTo ListVBProjectReferences_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Sub AuxGlobal.ListVBProjectReferences()"
    '----------------------------------------------------------------------------------------------------

30  Call CriaTabelaReferences
40  Set rsRefs = CurrentDb.OpenRecordset("tblVBAReferencias")
50  Call CurrentDb.Execute("DELETE FROM tblVBAReferencias WHERE maquina = '" & VBA.Environ("ComputerName") & "'")

60  Debug.Print VBA.String(120, "-")

70  For Each Ref In Access.Application.VBE.ActiveVBProject.References
80      If Ref.Type = 0 Then
90          Debug.Print "TYPE_LIB_DLL", "Ativo : " & Not CInt(Ref.IsBroken), Ref.Name, Ref.FullPath    ', ref.GUID
100     Else
110         Debug.Print "VBAPROJECT", "Ativo : " & Not CInt(Ref.IsBroken), VBA.Split(Ref.FullPath, "\")(UBound(VBA.Split(Ref.FullPath, "\"))), Ref.FullPath
120     End If

130     With rsRefs
140         rsRefs.addNew
150         !nomeRef = Ref.Name
160         !GUIDRef = Ref.GUID
170         !LocalArquivo = Ref.FullPath
180         !descricao = Ref.Description
190         !padrao = Ref.BuiltIn
200         !VersaoMajor = Ref.Major
210         !VersaoMinor = Ref.Minor
220         !Invalida = Ref.IsBroken
230         !TIPO = Ref.Type
240         !versaoAccess = Access.Application.Version
250         !buildAccess = Access.Application.Build
260         !data_hora = VBA.Now
270         !usuario = VBA.Environ("UserName")
280         !maquina = VBA.Environ("ComputerName")
290         .Update
300     End With
310 Next

Fim:
320 On Error GoTo 0
330 Exit Sub

ListVBProjectReferences_Error:
340 If VBA.Err <> 0 Then
350     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        'Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl, False, False)
360 End If
    GoTo Fim:
    'Debug Mode
370 Resume
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxGlobal.CriaTabelaReferences()
' TIPO             : Sub
' DATA/HORA        : 04/03/2015 12:14
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Recria a tabela que mantem a lista de referencias da ferramenta atualizada
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      04/03/2015 12:14    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Public Sub CriaTabelaReferences()
    Dim tb As Object, pSQLCreate As String
10  On Error Resume Next
20  Set tb = CurrentDb.TableDefs("tblVBAReferencias")

30  If tb Is Nothing Then

40      pSQLCreate = "CREATE TABLE tblVBAReferencias ( "
50      pSQLCreate = pSQLCreate & VBA.vbNewLine & "      nomeRef TEXT"
60      pSQLCreate = pSQLCreate & VBA.vbNewLine & "     ,GUIDRef TEXT"
70      pSQLCreate = pSQLCreate & VBA.vbNewLine & "     ,LocalArquivo TEXT"
80      pSQLCreate = pSQLCreate & VBA.vbNewLine & "     ,Descricao TEXT"
90      pSQLCreate = pSQLCreate & VBA.vbNewLine & "     ,Padrao YesNo"
100     pSQLCreate = pSQLCreate & VBA.vbNewLine & "     ,VersaoMajor INT"
110     pSQLCreate = pSQLCreate & VBA.vbNewLine & "     ,VersaoMinor INT"
120     pSQLCreate = pSQLCreate & VBA.vbNewLine & "     ,Invalida YesNo"
130     pSQLCreate = pSQLCreate & VBA.vbNewLine & "     ,tipo TEXT"
140     pSQLCreate = pSQLCreate & VBA.vbNewLine & "     ,versaoAccess TEXT"
150     pSQLCreate = pSQLCreate & VBA.vbNewLine & "     ,buildAccess TEXT"
160     pSQLCreate = pSQLCreate & VBA.vbNewLine & "     ,usuario TEXT"
170     pSQLCreate = pSQLCreate & VBA.vbNewLine & "     ,maquina TEXT"
180     pSQLCreate = pSQLCreate & VBA.vbNewLine & "     ,data_hora DateTime"

190     pSQLCreate = pSQLCreate & VBA.vbNewLine & ")"
200     On Error GoTo Erro
210     Call CurrentDb.Execute(pSQLCreate)
220 End If

Erro:
230 If VBA.Err <> 0 Then
240     VBA.MsgBox "Erro na criação da tabela de Referencias do Projeto VBA" & VBA.vbNewLine & VBA.vbNewLine & VBA.Error, VBA.vbCritical
250     Debug.Print pSQLCreate
260 End If

Fim:
270 On Error GoTo 0
280 Exit Sub
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxGlobal.ReInstallReferences()
' TIPO             : Sub
' DATA/HORA        : 04/03/2015 12:14
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Reinstala as referencias baseado na tabela de referencias atualizadas
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      04/03/2015 12:14    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Public Sub ReInstallReferences(Optional pComputerName As String)
    Dim tb     As Object
    Dim Ref    As Object
10  On Error GoTo ReInstallReferences_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Sub AuxGlobal.ReInstallReferences()"
    '----------------------------------------------------------------------------------------------------

    If pComputerName = "" Then pComputerName = Environ("ComputerName")

30  Set tb = CurrentDb.OpenRecordset("SELECT * FROM tblVBAReferencias WHERE Invalida = 0 and not Padrao = -1 and maquina = '" & pComputerName & "'")

40  Do While Not tb.EOF
        On Error Resume Next
        Set Ref = Nothing
50      Set Ref = Access.Application.References(tb!nomeRef.value)
60      If Ref Is Nothing Then
            On Error Resume Next
70          Set Ref = Access.Application.References.AddFromGuid(tb!GUIDRef.value, tb!VersaoMajor.value, tb!VersaoMinor.value)
            If VBA.Err = 0 Then
                Debug.Print "Referencia '" & tb!descricao.value & "' instalada com sucesso !"
            Else
                Debug.Print "Ocorreu o seguinte erro na instalação da Referencia '" & tb!descricao.value & "'/ Erro : " & VBA.Err.Description
            End If
80      Else
90          Debug.Print "Referencia '" & tb!descricao.value & "' ja instalada e ativa!  Local : " & tb!LocalArquivo.value
100     End If
110     tb.MoveNext
120 Loop

    Call ListVBProjectReferences
Fim:
130 On Error GoTo 0
140 Exit Sub

ReInstallReferences_Error:
150 If VBA.Err <> 0 Then
160     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        VBA.MsgBox strErrorMessagem, VBA.vbCritical
170     'Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl, False, False)
180 End If
    GoTo Fim:
    'Debug Mode
190 Resume
End Sub

Public Sub salvaVariavelAplicacao(pName As String, pValue As String)
    Dim rsAux As Object, pUserTable As String
    On Error GoTo saveConfig_Error
    Set rsAux = Access.CurrentDb.OpenRecordset("tblApp")
    rsAux.edit
    rsAux.Fields(pName).value = pValue
    rsAux.Update
    rsAux.Close
    Call Publicas.RemoverObjetosMemoria(rsAux)
    On Error GoTo 0
    Exit Sub
saveConfig_Error:
    If VBA.Err <> 0 Then Call TratarErro(VBA.Err.Description, VBA.Err.Number, "AuxAplicacao.salvaValor()", VBA.Erl, False, False)
    Exit Sub
    Resume
End Sub

Public Function pegarVariavelAplicacao(pName As String)
    pegarVariavelAplicacao = Nz(DLookup(pName, "tblApp"), "")
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxGlobal.ConfigAppImages()
' TIPO             : Sub
' DATA/HORA        : 03/09/2014 11:37
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Configura as imagens da aplicação
'---------------------------------------------------------------------------------------
Public Sub ConfigAppImages()
    '    On Error Resume Next
    Dim strEndereco As String
    Dim dirPath As String
10  On Error GoTo ConfiguraIconeAplicativo_Error

20  dirPath = getDBPath & "\Settings"
30  strEndereco = dirPath & "\" & Nz(DLookup("filename", "Pegar_NomeArquivoTemplate", "NomeTemplate='icon_app'"))

40  If Not AuxFileSystem.FileExists(strEndereco) Then
50      strEndereco = AuxTabela.PegaCaminhoArquivoAnexo("icon_app", "tblAnexos", dirPath, "OriginalName", False)
60  End If

70  strEndereco = AuxFileSystem.getUNCFullName(strEndereco)
80  Call salvaVariavelAplicacao("icon_file_path", strEndereco)

90  Call DefinirPropriedadeObjeto(CurrentDb, "UseAppIconForFrmRpt", -1, 1, False)
100 Call DefinirPropriedadeObjeto(CurrentDb, "AppIcon", strEndereco, 10, False)

110 strEndereco = AuxTabela.PegaCaminhoArquivoAnexo("logo_app", "tblAnexos", dirPath, "OriginalName", False)
120 strEndereco = AuxFileSystem.getUNCFullName(strEndereco)
130 Call salvaVariavelAplicacao("logo_file_path", strEndereco)

140 On Error GoTo 0
150 Exit Sub

ConfiguraIconeAplicativo_Error:
160 If VBA.Err <> 0 Then
170     Call RegistraLog("AuxGlobal.ConfiguraIconeAplicativo() - Linha : " & VBA.Erl & " / ERROR : " & VBA.Err.Number & "-" & VBA.Err.Description, CurrentProject.Path & "\Log\VBAErros.txt")
180 End If
190 Exit Sub
200 Resume
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxForm.DisplayImage()
' TIPO             : Function
' DATA/HORA        : 10/02/2015 23:17
' CONSULTOR        : (Adelson)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Carrega uma imagem (gravada em disco) em um controle de imagem
'---------------------------------------------------------------------------------------
Public Function DisplayImage(ctlImageControl As Control, strImagePath As Variant) As String
    '----------------------------------------------------------------------------------------------------
    On Error GoTo DisplayImage_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
    Dim dtSartRunProc As Date: dtSartRunProc = Time
    Dim varivaeis As Variant
    Const cstr_ProcedureName As String = "Function prjBRAMFOF.Módulo1.DisplayImage()"
    '----------------------------------------------------------------------------------------------------

    Dim strResult As String
    Dim strDatabasePath As String
    Dim intSlashLocation As Integer

    With ctlImageControl
        If VBA.IsNull(strImagePath) Then
            .Visible = False
            strResult = "No image name specified."
        Else
            'Endereço Relativo
            If InStr(1, strImagePath, "\") = 0 Then
                strDatabasePath = CurrentProject.FullName
                intSlashLocation = VBA.InStrRev(strDatabasePath, "\", Len(strDatabasePath))
                strDatabasePath = VBA.Left(strDatabasePath, intSlashLocation)
                strImagePath = strDatabasePath & strImagePath
            End If
            .Visible = True
            .Picture = strImagePath
            strResult = "Image found and displayed."
        End If
    End With

Exit_DisplayImage:
        DisplayImage = strResult
        Exit Function

    On Error GoTo 0
    Exit Function

DisplayImage_Error:
        Select Case VBA.Err.Number
            Case 2220       ' Can't find the picture.
                ctlImageControl.Visible = False
                strResult = "Can't find image in the specified name."
                Resume Exit_DisplayImage:
            Case Else       ' Some other error.
                lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
                varivaeis = Array(vbNullString)
                'Tratamento de erro personalizado
                Call TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl, False, False)
                Resume Exit_DisplayImage:
        End Select

    Exit Function
    'Debug Mode
    Resume
End Function

'---------------------------------------------------------------------------------------
' Procedure : SalvaVariavelLocal
' Author    : Adelson
' Date      : 06/12/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'''---------------------------------------------------------------------------------------
''' PROCEDIMENTO     : SalvaVariavelLocal()
''' TIPO             : Sub
''' DATA/HORA        : 06/12/2013 14:20
''' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
''' CONTATO          : adelson@tecnun.com.br
''' DESCRIÇÃO        : Salva uma variável do usuário
'''---------------------------------------------------------------------------------------
''Sub SalvaVariavelLocal(pName As String, pValue As String)
''    Dim rsAux As Object, pUserTable As String
''    On Error GoTo SalvaVariavelLocal_Error
''
''    'pUserTable = CriaTabelaVariaveisUsuario()
''
''    Set rsAux = CodeDb.OpenRecordset("SELECT ConfigName, configValue FROM [" & pUserTable & "] WHERE configName = '" & pName & "'")
''    If Not rsAux Is Nothing Then
''        If rsAux.EOF Then rsAux.AddNew Else rsAux.Edit
''        rsAux!ConfigName = pName
''        rsAux!configValue = pValue
''        rsAux.Update
''        rsAux.Close
''    End If
''    Set rsAux = Nothing
''
''    On Error GoTo 0
''    Exit Sub
''
''SalvaVariavelLocal_Error:
''    If VBA.Err <> 0 Then Call auxTFW.TratarErro(VBA.Err.Description, VBA.Err.Number, "SalvaVariavelLocal()", VBA.Erl)
''    Exit Sub
''    Resume
''End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : pegaValor()
' TIPO             : Function
' DATA/HORA        : 06/12/2013 14:21
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Recupera uma variável do usuario
'---------------------------------------------------------------------------------------
Function pegaValor(pNomeVariavel As String, Optional valorSeNull, Optional eEscopo As eEscopoVariavel = varglobal) As String
    On Error GoTo pegaValor_Error
    Dim rsAux As Object
    With CodeDb.QueryDefs("Pegar_ValorVariavel")
        .Parameters("@nome").value = pNomeVariavel
        .Parameters("@Usuario").value = VBA.IIf(eEscopo = varUsuario, ChaveUsuario, Null)
        Set rsAux = .OpenRecordset()
    End With
    If Not rsAux Is Nothing Then If Not rsAux.EOF Then pegaValor = Nz(rsAux!ValorVariavel.value, valorSeNull): rsAux.Close
    Set rsAux = Nothing
    If Not VBA.IsMissing(valorSeNull) Then
        If pegaValor = "" Then pegaValor = valorSeNull
    End If
    On Error GoTo 0
    Exit Function
pegaValor_Error:
    If VBA.Err <> 0 Then
        Call TratarErro(VBA.Err.Description, VBA.Err.Number, "pegaValor()", VBA.Erl, False, False)
    End If
    Exit Function
    Resume
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : salvaValor()
' TIPO             : Sub
' DATA/HORA        : 29/04/2014 14:21
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Salva uma variável global da aplicação
'---------------------------------------------------------------------------------------

Function salvaValor(pNomeVariavel As String, pValorVariavel As Variant, Optional eEscopo As eEscopoVariavel = varglobal) As Boolean
    Dim rsAux As Object
10  On Error GoTo salvaValor_Error

    Call LimparVariaveisSessao(pNomeVariavel)
    If VBA.IsMissing(pValorVariavel) Then Exit Function
    If pValorVariavel = "" Then Exit Function
    
    With CurrentDb.QueryDefs("Insere_Variavel")
        .Parameters("@nomeVariavel").value = pNomeVariavel
        .Parameters("@valorVariavel").value = pValorVariavel
        .Parameters("@usuario").value = VBA.IIf(eEscopo = varUsuario, ChaveUsuario, Null)
        Call .Execute
    End With
    
100 salvaValor = VBA.Err.Number = 0
110 On Error GoTo 0
120 Exit Function

salvaValor_Error:
130 If VBA.Err <> 0 Then Call TratarErro(VBA.Err.Description, VBA.Err.Number, "salvaValor()", VBA.Erl)
140 Exit Function
150 Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : pegaValor()
' TIPO             : Function
' DATA/HORA        : 06/12/2013 14:21
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Recupera uma variável do usuario
'---------------------------------------------------------------------------------------
Function pegarVariavel(Optional pNomeVariavel As String)
    On Error GoTo pegaValor_Error
    Dim rsAux As Object
    If Conexao Is Nothing Then Inicializar_Globais
    Set rsAux = Conexao.PegarRS("Pegar_Variavel", VBA.IIf(pNomeVariavel <> "", pNomeVariavel, Null))
    If Not rsAux Is Nothing Then If Not rsAux.EOF Then Set pegarVariavel = rsAux
    Set rsAux = Nothing
    On Error GoTo 0
    Exit Function
pegaValor_Error:
    If VBA.Err <> 0 Then Call TratarErro(VBA.Err.Description, VBA.Err.Number, "pegarVariavel()", VBA.Erl, False, False)
    Exit Function
    Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : LimparVariaveisSessao()
' TIPO             : Function
' DATA/HORA        : 09/06/2015 23:21
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Limpa uma ou mais variáveis da tabela de sessão do usuario
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      09/06/2015 23:21    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Function LimparVariaveisSessao(Optional pNomeVariavel As String)
10  On Error GoTo LimparVariaveisSessao_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function LimparVariaveisSessao()"
    '----------------------------------------------------------------------------------------------------

30  With CodeDb.QueryDefs("Limpar_Variaveis")
40      .Parameters("@nomeVariavel").value = pNomeVariavel
50      .Parameters("@Usuario").value = ChaveUsuario
60      Call .Execute
70  End With

Fim:
80  On Error GoTo 0
90  Exit Function

LimparVariaveisSessao_Error:
100 If VBA.Err <> 0 Then
110     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
120     Call TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
130 End If
    GoTo Fim:
    'Debug Mode
140 Resume
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : IDSession()
' TIPO             : Function
' DATA/HORA        : 29/04/2014 14:54
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Gera ou recupera uma nova IDSession do usuário local
'---------------------------------------------------------------------------------------
Function IDSession(Optional bNew As Boolean) As String
    On Error GoTo IDSession_Error
    If bNew Then
        UserIDSession = VBA.Format(VBA.Now, "yyyymmddhhnnss")
        Call salvaValor("IDSession", UserIDSession)
    Else
        UserIDSession = pegaValor("IDSession")
    End If
    IDSession = UserIDSession
    On Error GoTo 0
    Exit Function

IDSession_Error:
    If VBA.Err <> 0 Then Call TratarErro(VBA.Err.Description, VBA.Err.Number, "IDSession()", VBA.Erl)
    Exit Function
    Resume
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : InicializaSessaoUsuario()
' TIPO             : Sub
' DATA/HORA        : 09/06/2015 23:17
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Inicializa variáveis de uma nova sessão do usuário
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      09/06/2015 23:17    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Sub InicializaSessaoUsuario()
    Dim rs As Object
    On Error GoTo InicializaSessaoUsuario_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
    Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Sub InicializaSessaoUsuario()"
    '----------------------------------------------------------------------------------------------------

    On Error GoTo InicializaSessaoUsuario_Error
    Dim dbs
    Dim i      As Integer
    Dim maxCaract As Long

    Call FinalizarAplicação(bLog:=False, bClearGlobal:=False)
    
    'maxCaract = Nz(CodeDb.OpenRecordset("Select Max(QtdCaract) from PegarListaArquivosVinculados").Fields(0).value, 0)

    'Registra o acesso do usuário
    'Call adicionaUsuarioGrupo
    'Finaliza antes de criar novas variáveis
    'Cria novamente as variáveis com novo ID

''    Call auxAuditoria.RegistraLog(String(150, "*"))
''    Call auxAuditoria.RegistraLog("Sessão iniciada por '" & ChaveUsuario & "' ás '" & VBA.Now) ' & "' usando a IDSession : " & UserIDSession)
''    Call auxAuditoria.RegistraLog(" # DATA BASES")
''  Call auxAuditoria.RegistraLog("     > " & VBA.Left("CLIENT " & VBA.Space(11), 11) & " : " & VBA.Left(CodeDb.name & VBA.Space(maxCaract), maxCaract + 4) & formatFileSize(VBA.FileLen(CodeDb.name), True) & VBA.Space(2) & VBA.FileDateTime(CodeDb.name))

''    Set rs = CodeDb.QueryDefs("PegarListaArquivosVinculados").OpenRecordset
''
''    If rs.RecordCount > 0 Then
''        dbs = rs.GetRows(rs.RecordCount)
''        For i = 0 To UBound(dbs, 2)
''            If Not dbs(0, i) Like "*https://*" Then
''            If VBA.Dir(CStr(dbs(0, i))) <> "" Then
''                Call auxAuditoria.RegistraLog("     > BACK-END : " & VBA.Left(dbs(0, i) & VBA.Space(maxCaract), maxCaract + 4) & VBA.Left(auxAuditoria.formatFileSize(FileSystem.FileLen(CStr(dbs(0, i)))) & VBA.Space(14), 14) & auxAuditoria.FileDateTime(CStr(dbs(0, i))))
''            Else
''                Call auxAuditoria.RegistraLog("     > BACK-END : " & VBA.Left(dbs(0, i) & VBA.Space(maxCaract), maxCaract + 4) & "(Arquivo nao localizado)")
''            End If
''            End If
''        Next
''    End If

    'Call auxAuditoria.RegistraLog(String(150, "*"))
    Call AuxAplicacao.salvaValor("DataHoraLogon", VBA.Now)

Fim:
    On Error GoTo 0
    Exit Sub

InicializaSessaoUsuario_Error:
    If VBA.Err <> 0 Then
        lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        Call TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
    End If
    GoTo Fim:
    'Debug Mode
    Resume

End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : FinalizarAplicação()
' TIPO             : Sub
' DATA/HORA        : 29/04/2014 15:38
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Finaliza a sessão do usuário mantendo salvos algumas  variáveis
'---------------------------------------------------------------------------------------
Public Sub FinalizarAplicação(Optional bLog As Boolean = True, Optional bClearGlobal As Boolean = True)
    Dim pUserTable As String
    Dim ultima_selecao() As Variant
    On Error GoTo FinalizarAplicação_Error

    'Apaga a tabela temporária
    Call RegistraLog("Remove links das tabelas remotas do usuário....")
'    Call configurarViewDados(Remover:=True)
    If bLog Then
        Call RegistraLog(String(150, "*"))
        Call RegistraLog("A Sessão '" & UserIDSession & " foi finalizada por '" & ChaveUsuario & "' ás '" & VBA.Now & "'")
        Call RegistraLog(String(5, "-"))
    End If
    On Error GoTo 0
    Exit Sub

FinalizarAplicação_Error:
    If VBA.Err <> 0 Then Call TratarErro(VBA.Err.Description, VBA.Err.Number, "FinalizarAplicação()", VBA.Erl)
    Exit Sub
    Resume

End Sub

Function PegarHoraLogon(): PegarHoraLogon = pegaValor("DataHoraLogon"): End Function
Function PegarNomeFerramenta(): PegarNomeFerramenta = pegarVariavelAplicacao("appName"): End Function
Function PegarDescricaoFerramenta(): PegarDescricaoFerramenta = pegarVariavelAplicacao("appDesc"): End Function




Sub FinalizaAccess()
    Access.Application.CloseCurrentDatabase
    Access.Application.quit
End Sub


'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxAplicacao.CriarNovaAplicacao()
' TIPO             : Sub
' DATA/HORA        : 15/06/2015 15:20
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Copia os arquivos da Aplicação para um novo local para que seja configurado
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      15/06/2015 15:20    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Sub CriarNovaAplicacao(sAppName As String, sCustomer As String, icon_path As String, logo_path As String, img_bg_inico As String, color As Long)

10  On Error GoTo CriarNovaAplicacao_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Sub AuxAplicacao.CriarNovaAplicacao()"
    '----------------------------------------------------------------------------------------------------
    Dim strNovaApp As String
    Dim strLocal As String
    Dim bResult As Boolean
    Dim db As Object
    Dim sPatasDestino As String
    Dim arquivo_cliente As String
    Dim rs As Object

30  strNovaApp = sAppName

40  strLocal = CaixaDeDialogo(msoFileDialogSaveAs, "Selecione o local onde será salvo os arquivos")
50  strLocal = strLocal & "\" & strNovaApp & "_v" & VBA.Format(Date, "yyyy.mm.dd.Rev.001")
60  Call MkFullDirectory(strLocal)
70  sPatasDestino = AuxFileSystem.getFolderSH(CurrentProject.FullName).ParentFolder.Path

80  Call AuxFileSystem.CopiarPastaPara(origem:=sPatasDestino, destino:=strLocal)

90  arquivo_cliente = strLocal & "\Programa\" & strNovaApp & "_v1.00.0.accdb"

    '100 If bResult Then
100 Name strLocal & "\Programa\" & CurrentProject.Name As arquivo_cliente
110 Call VBA.Kill(strLocal & "\Programa\" & VBA.Replace(CurrentProject.Name, "accdb", "laccdb"))

120 Set db = DBEngine.OpenDatabase(arquivo_cliente)

130 Set rs = db.OpenRecordset("tblApp")
140 If rs.EOF Then rs.addNew Else rs.edit
150 rs!appName.value = strNovaApp
160 rs!CustomerName.value = sCustomer
170 rs!icon_file_path.value = icon_path
180 rs!logo_file_path.value = logo_path
190 rs!img_bg_inico.value = img_bg_inico
200 rs!color_bar.value = color
210 rs.Update

220 Call db.Execute("DELETE FROM tblVersao")
230 Set rs = db.OpenRecordset("tblVersao")
310 With rs
        .addNew
320     !versao.value = "1.00.0"
330     !dataCriacao.value = VBA.Now
340     !usuario.value = ChaveUsuario
350     !maquina.value = Environ("ComputerName")
360     !Anotacoes.value = "Versão Inciada a partir do TECUN TFW App Creator"
370     .Update
380 End With

    'Remove o botão para criação de nova Aplicação.
    Call db.Execute("DELETE FROM tblRibbon_Controls WHERE ID ='btnNovaAplicacao'")


390 Call MessageBoxMaster("F021")
400 Call FinalizaAccess

Fim:
410 On Error GoTo 0
420 Exit Sub

CriarNovaAplicacao_Error:
430 If VBA.Err <> 0 Then
440     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
450     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
460 End If
    GoTo Fim:
    'Debug Mode
470 Resume
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : AuxAplicacao.salvarVersaoAplicacao()
' TIPO             : Sub
' DATA/HORA        : 15/06/2015 18:01
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Altera informações sobre a versão
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      15/06/2015 18:01    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Function salvarVersaoAplicacao() As Boolean
    Dim vNovaVersao As String, sAnotacoes As String
    Dim rsVersao As Object, vPartVer, sTipoErro As String, bVersaoOk As Boolean
1   On Error GoTo salvarVersaoAplicacao_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
2   Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Sub AuxAplicacao.salvarVersaoAplicacao()"
    '----------------------------------------------------------------------------------------------------
    Dim bAlterou As Boolean

3   vPartVer = VBA.Split(AuxAplicacao.NumeroVersaoAtual, ".")

4   vNovaVersao = VBA.InputBox("Informe a versão para a aplicação no formato X.XX.N" & VBA.vbNewLine & _
                           "O ultimo caractere N deve ser : " & VBA.vbNewLine & _
                         "    0 - Desenvolvimento" & VBA.vbNewLine & _
                         "    1 - Homologação" & VBA.vbNewLine & _
                         "    2 - Produção", , vPartVer(0) & "." & VBA.Format(VBA.Val(vPartVer(1)) + 1, "00") & "." & vPartVer(2))

5   If vNovaVersao = "" Then GoTo Fim

6   vPartVer = VBA.Split(vNovaVersao, ".")
7   bVersaoOk = True
8   If UBound(vPartVer) <> 2 Then bVersaoOk = False: sTipoErro = "Erro Formato": GoTo ValidaVersao
9   If Not VBA.IsNumeric(vPartVer(0)) And Not VBA.IsNumeric(vPartVer(1)) And Not VBA.IsNumeric(vPartVer(2)) Then bVersaoOk = False: sTipoErro = "Não numericos": GoTo ValidaVersao
10  If Not (vPartVer(2) >= 0 And vPartVer(2) <= 2) Then bVersaoOk = False: sTipoErro = "Caractere Ambiente Prod. Desenv, Homolog": GoTo ValidaVersao

ValidaVersao:
11  If Not bVersaoOk Then
12      VBA.MsgBox "Versão '" & vNovaVersao & " inválida !" & VBA.vbNewLine & VBA.vbNewLine & "Motivo : " & sTipoErro, VBA.vbExclamation
13      GoTo Fim
14  End If


15  If vNovaVersao <> "" Then
16      sAnotacoes = VBA.InputBox("Informa uma descrição sobre a versão !")
17      Set rsVersao = CurrentDb.OpenRecordset("SELECT * FROM tblVersao WHERE Versao = '" & vNovaVersao & "'")
18      If Not rsVersao.EOF Then
19          rsVersao.edit
20          bAlterou = rsVersao!versao.value <> vNovaVersao
21      Else
22          rsVersao.addNew
23          bAlterou = True
24      End If
25      With rsVersao
26          !versao.value = vNovaVersao
27          !dataCriacao.value = VBA.Now
28          !usuario.value = ChaveUsuario
29          !maquina.value = Environ("ComputerName")
30          !Anotacoes.value = sAnotacoes
            '200         !idAplicacao.Value = CurrentProject.Name
31          .Update
32      End With
33  End If

34  salvarVersaoAplicacao = bAlterou

Fim:
35  On Error GoTo 0
36  Exit Function

salvarVersaoAplicacao_Error:
37  If VBA.Err <> 0 Then
38      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
39      Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
40  End If
    GoTo Fim:
    'Debug Mode
41  Resume
End Function


Function NumeroVersaoAtual() As String
     NumeroVersaoAtual = CurrentDb.OpenRecordset("Pegar_Versao_Atual")(0).value
End Function

Function NomeVersaoAtual() As String
    NomeVersaoAtual = Publicas.Carregar_Versao
End Function

Function LocalReferenciaTFW() As String
    On Error Resume Next
    LocalReferenciaTFW = Access.References!tfw.FullPath
    If VBA.Err <> 0 Then
        LocalReferenciaTFW = "(Referencia do TFW não instalada !)"
    End If
End Function

Function PegaEnderecoBDs() As String
    PegaEnderecoBDs = PegaEnderecoCliente & "\BDs"
End Function

Function PegaEnderecoCliente() As String
    PegaEnderecoCliente = PegarPasta(CodeProject.Path)
End Function

Function PegaEnderecoDBsDados() As String
    PegaEnderecoDBsDados = PegaEnderecoBDs & "\Dados"
End Function

Function PegaEnderecoConfiguracoes() As String
    PegaEnderecoConfiguracoes = PegaEnderecoBDs & "\Configuracoes"
End Function

Function PegaEndereco_Programa() As String
    PegaEndereco_Programa = FormatPath(CurrentProject.Path)
End Function

Function PegaEndereco_Templates() As String
    PegaEndereco_Templates = PegaEndereco_Programa & "\Templates"
    MkFullDirectory PegaEndereco_Templates
End Function

Function PegaEndereco_Relatorios() As String
    PegaEndereco_Relatorios = PegaEndereco_Programa & "\Relatorios"
    MkFullDirectory PegaEndereco_Relatorios
End Function

Function DevMode() As Boolean
    DevMode = VBA.Left(VBA.Split(NomeVersaoAtual, ".")(2), 1) = "0"
End Function

Public Property Get LoginWindows() As String
    m_sChaveUsuario = VBA.Environ("Username")
    m_sChaveUsuario = VBA.Replace(m_sChaveUsuario, " ", "")
    LoginWindows = UCase(m_sChaveUsuario)
End Property

Sub ModoProducao(bModoProducao As Boolean)
    If bModoProducao Then
        Call SetStartupOptions("ModoProducao", DB_BOOLEAN, False)
    Else
        Call SetStartupOptions("ModoProducao", DB_BOOLEAN, True)
    End If
    MsgBox "Necessário reiniciar o aplicativo para habilitar o modo de Execução (Produção/Homologação ou Desenvolvimento)", VBA.vbExclamation, "Versão"
'    '----------------------------------------------------
'    'Configura o aplicativo no modo produção
'    '----------------------------------------------------
'    Call DefinirPropriedadeObjeto(CurrentDb, "ShowDocumentTabs", Not bModoProducao, 1, False)
'    Call DefinirPropriedadeObjeto(CurrentDb, "AllowBreakIntoCode", Not bModoProducao, 1, False)
'    Call DefinirPropriedadeObjeto(CurrentDb, "AllowSpecialKeys", Not bModoProducao, 1, False)
'    Call DefinirPropriedadeObjeto(CurrentDb, "AllowBypassKey", Not bModoProducao, 1, False)
'    Call DefinirPropriedadeObjeto(CurrentDb, "AllowFullMenus", Not bModoProducao, 1, False)
'    Call DefinirPropriedadeObjeto(CurrentDb, "StartUpShowDBWindow", Not bModoProducao, 1, False)
    
End Sub
