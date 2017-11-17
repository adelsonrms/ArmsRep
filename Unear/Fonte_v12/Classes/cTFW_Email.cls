VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTFW_Email"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event OnLog(pMessage As String)
Public Event OnProgress(lngMax As Long, pMessage As String)

Dim Mensagem As String
Private Const MSOffice_Outlook As String = "Outlook.Application"

Private objOutlook As Object
Private OLDialog As Object
Private m_sAssunto As String
Private m_sCC As String
Private m_sCCO As String
Private m_sOnBehalfOf As String
Private m_sPara As String
Private m_vBody As Variant

Public Enum eMailActions
    Email_Display = 1
    Email_Send = 2
    Email_SalvarRascunho = 3
End Enum

Public Enum eEmailCopyFormat
    EmailCopyFormat_HTML = 0
    EmailCopyFormat_IMAGE = 1
    EmailCopyFormat_TEXT = 2
    EmailCopyFormat_TABLE = 3
End Enum

Private p_oMail As Object

Private m_objVariaveisCorpo As Collection
Private m_objVariaveisAssunto As Collection

Private m_objMarcadoresCorpo As Collection
Private m_oValidacaoDestinatarios As Variant

Private m_objSession As Object
Private m_colEmbeddedImages As Collection

Private Sub Class_Initialize()
    Set m_objVariaveisCorpo = New Collection
    Set m_colEmbeddedImages = New Collection
End Sub

Private Sub Class_Terminate()
    Set p_oMail = Nothing
    Set m_objVariaveisCorpo = Nothing
    Set m_objSession = Nothing
    Set m_colEmbeddedImages = Nothing
End Sub
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cEmail.StartOutlook()
' TIPO             : Function
' DATA/HORA        : 04/07/2016 10:28
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Inicializa uma nova instancia do Outlook
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      04/07/2016 10:28
'---------------------------------------------------------------------------------------
Private Function StartOutlook(Optional bNew As Boolean) As Object
'Inicializa o Outlook
'---------------------------------------------------------------------------------------
1   On Error GoTo StartOutlook_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
2   Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function cEmail.StartOutlook()"
    '---------------------------------------------------------------------------------------
3   On Error Resume Next
    'Cria/Recupera a instancia do outlook
4   If Not bNew Then Set objOutlook = VBA.GetObject(Class:=MSOffice_Outlook)
5   If objOutlook Is Nothing Then Set objOutlook = VBA.CreateObject(Class:=MSOffice_Outlook)



6   Set Me.Session = objOutlook.Session
7   Set StartOutlook = objOutlook

Fim:
8   On Error GoTo 0
9   Exit Function

StartOutlook_Error:
10  If Err <> 0 Then
11      lngErrorNumber = VBA.Err.Number: strErrorMessagem = "Erro na instalação do Outlook " & vbNewLine & vbNewLine & VBA.Err.Description
        Debug.Print cstr_ProcedureName, "Linha : " & VBA.Erl() & " - " & strErrorMessagem
12      'Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl())
13  End If
    GoTo Fim:
    'Debug Mode
14  Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cEmail.CriarNovoEmail()
' TIPO             : Function
' DATA/HORA        : 04/07/2016 10:37
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Criar um novo email ou carrega uma mensagem salva no disco
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      04/07/2016 10:37
'---------------------------------------------------------------------------------------
Function CriarNovoEmail(Optional pFromFile As String) As Object
    Dim oEmail As Object
    '---------------------------------------------------------------------------------------
1   On Error GoTo CriarNovoEmail_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
2   Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function cEmail.CriarNovoEmail()"
    '---------------------------------------------------------------------------------------
    
    Call StartOutlook
    
3   If pFromFile <> "" Then
4       Set p_oMail = objOutlook.CreateItemFromTemplate(pFromFile)
5   Else
6       Set p_oMail = objOutlook.CreateItem(0)
7   End If

8   Set CriarNovoEmail = p_oMail

Fim:
9   On Error GoTo 0
10  Exit Function

CriarNovoEmail_Error:
11  If VBA.Err <> 0 Then
12      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        'Debug.Print cstr_ProcedureName, "Linha : " & VBA.Erl() & " - " & strErrorMessagem
13      'Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl())
14  End If
    GoTo Fim:
    'Debug Mode
15  Resume
End Function

Function MailFormats() As Variant
    MailFormats = Array("HTML", "IMAGE", "TEXT", "TABLE")
End Function

Function MailActions() As Variant
    MailActions = Array("Send", "Display")
End Function
'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cEmail.CriarSubPastaNaPastaDeEntrada()
' TIPO             : Function
' DATA/HORA        : 04/07/2016 10:23
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Cria uma subpasta na pasta de entrada
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      04/07/2016 10:23
'---------------------------------------------------------------------------------------
Function CriarSubPastaNaPastaDeEntrada(pFolder As String) As Object
    Dim inbox As Object, oSubFolder As Object
    'olFolderInbox = 6
    '---------------------------------------------------------------------------------------
1   On Error GoTo CriarSubPastaNaPastaDeEntrada_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
2   Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function cEmail.CriarSubPastaNaPastaDeEntrada()"
    '---------------------------------------------------------------------------------------
3   Set inbox = Me.Session.GetDefaultFolder(6)
4   On Error Resume Next
5   Set oSubFolder = inbox.Folders(pFolder)
6   If Err <> 0 Then Set oSubFolder = inbox.Folders.Add(pFolder)
7   Set CriarSubPastaNaPastaDeEntrada = oSubFolder
Fim:
8   On Error GoTo 0
9   Exit Function

CriarSubPastaNaPastaDeEntrada_Error:
10  If Err <> 0 Then
11      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        'Debug.Print cstr_ProcedureName, "Linha : " & VBA.Erl() & " - " & strErrorMessagem
12      'Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl())
13  End If
    GoTo Fim:
    'Debug Mode
14  Resume
End Function

Public Property Get EmbeddedImages() As Collection
    Set EmbeddedImages = m_colEmbeddedImages
End Property

Public Property Set EmbeddedImages(colEmbeddedImages As Collection)
    Set m_colEmbeddedImages = colEmbeddedImages
End Property

Sub Display()
    Call Me.Enviar(Me.Body, Email_Display)
End Sub

Sub SalvarRascunho()
    Call Me.Enviar(Me.Body, Email_SalvarRascunho)
End Sub

'--------------------------------------------------------------------------------------------------------------------------------
'Autor      : Adelson Silva - Mondial Informatica (adelsons@gmail.com)
'Data/Hora  : 04/09/2013 15:42
'Descrição  : Criar/Envia um email usando o Outlook com diferentes formatos no corpo do email
'             Prepara o arquivo e depois o importa
'             O argumento pRangeCopy pode ser passado para copiar uma seleção em uma planilha qualquer para o email
'             Para indicar o formato, usar o argumento pPasteAs para colar o conteuodo do pRangeCopy com as seguinte opções :
'              IMG ou IMAGE ou I    : Colar como imagem
'              TXT ou TEXT ou T     : Colar como texto
'              TBL ou TABLE         : Colar o intervalo como tabela padrão do Excel
'              HTM ou HTML ou H     : Salva o intervalo em um arquivo temporário html, obtem o codigo fonte HTML salva como corpo do email
'---------------------------------------------------------------------------------------------------------------------------------
'27/08/2013 - Criado a função
'04/06/2013 - Inclusão de argumentos para envio em formato HTML, Imagem e Texto
'---------------------------------------------------------------------------------------------------------------------------------
Function Enviar(Optional pBody, _
              Optional strAction As eMailActions = eMailActions.Email_Display, _
              Optional strAttachmentPaths As String, _
              Optional pPasteAs As String = "Image", Optional bCloseEmail As Boolean = True) As Variant
    
    Dim wdDoc               As Object
    Dim wb                  As Object
    Dim sTempFile           As String
    Dim sTempFileSource     As String
    Dim pRangeCopy          As Object


CreateEmail:
10  On Error GoTo Err_Outlook
    'Cria um novo email
    If p_oMail Is Nothing Then Set p_oMail = Me.CriarNovoEmail()

40  On Error GoTo 0

    If Not VBA.IsMissing(pBody) Then
        If VBA.IsObject(pBody) Then
            Set m_vBody = pBody
        Else
            m_vBody = pBody
        End If
    End If
    
    'Criar e Envia o email
50  With p_oMail
        'Trata o Assunto
        If .subject <> "" Then m_sAssunto = .subject
        If Me.VariaveisAssunto.count > 0 Then m_sAssunto = PreencherVariaveis(m_sAssunto, Me.VariaveisAssunto)
        If m_sAssunto <> "" Then .subject = m_sAssunto

70      .SentOnBehalfOfName = m_sOnBehalfOf
        'Anexos
80      If strAttachmentPaths <> "" Then
90          If strAttachmentPaths <> "attach tmp" Then
100             Call FillMailItems(strAttachmentPaths, .Attachments)
110         End If
120     End If
130     'Endereçamento
        .To = m_sPara
140     If m_sCC <> "" Then .cc = m_sCC
150     If m_sCCO <> "" Then .BCC = m_sCCO
        'Preenche as VariaveisCorpo/imagens
        'Assunto
160     If Not VBA.IsObject(m_vBody) Then
            'Trata variáveis no Corpo
            If Me.MarcadoresCorpo.count > 0 Then Call PreencherMarcadores(Me.MarcadoresCorpo)
            If m_vBody <> "" Then
                If Me.VariaveisCorpo.count > 0 Then m_vBody = PreencherVariaveis(VBA.CStr(m_vBody), Me.VariaveisCorpo)
                .HTMLBody = m_vBody
            End If
180     Else
190         If VBA.TypeName(m_vBody) = "Range" Then
200             Set pRangeCopy = m_vBody
210             pPasteAs = UCase(pPasteAs)
220             If pPasteAs = "HTML" Or pPasteAs = "HTM" Or pPasteAs = "H" Then
                    'Se for html precisamos obter o fonte html do endereço. Pra isso temos que salva o arquivo em um local temporário
230                 Set wb = pRangeCopy.Application.Workbooks.Add(1)
240                 pRangeCopy.Copy
250                 With wb.Worksheets(1).Range("A1")
260                     .PasteSpecial -4163    'xlPasteValues
270                     .PasteSpecial -4122    'xlPasteFormats
280                     .PasteSpecial 8    'xlPasteColumnWidths
290                 End With
300                 wb.Worksheets(1).usedrange.entirecolumn.AutoFit
310                 pRangeCopy.Application.CutCopyMode = False
320                 sTempFile = Environ("temp") & "\" & VBA.Replace(VBA.Format(VBA.Now, "0.000000000"), ".", "") & ".html"
330                 sTempFileSource = VBA.Replace(sTempFile, ".html", ".xlsx")
340                 Call wb.SaveAs(sTempFileSource)
350                 If sTempFile <> "" Then If VBA.Dir(sTempFile) <> "" Then VBA.Kill sTempFile
                    'xlSourceRange = 4
                    'xlHtmlStatic = 0
360                 wb.PublishObjects.Add(4, sTempFile, wb.Worksheets(1).Name, wb.Worksheets(1).usedrange.Address, 0).Publish (True)
370                 wb.Close False
380                 .HTMLBody = ReadTextFile(sTempFile)
                    'Alinha o conteudo HTML a esquerda (O excel salva uma planilha publicada à direita)
390                 .HTMLBody = VBA.Replace(.HTMLBody, "align=center x:publishsource=", "align=left x:publishsource=")
400                 If strAttachmentPaths = "attach tmp" Then Call FillMailItems(sTempFileSource, .Attachments)
410                 .Display
420             Else
                    'wdDoc : Instancia de um Document Word
                    'Recupera a instancia do Documento Word que representa o email
430                 Set wdDoc = .GetInspector.WordEditor
                    'Algumas funcionalidades são pertencentes ao documento. Com o acesso a instancia podemos executar as ações
                    'como se estivessemos em um documento
440                 pRangeCopy.Copy
                    'Pra conseguirmos colar precisamos exibir o email antes
450                 .Display
                    'Na sequencia colamos o conteudo
                    'Esse metodo cola como imagem atraves da constante
                    'wdFormatPlainText  = 22 (Colar como Texto)
                    'wdChartPicture     = 13 (Colar como imagem)
460                 If pPasteAs = "IMAGE" Or pPasteAs = "IMG" Or pPasteAs = "I" Then
470                     wdDoc.Application.Selection.PasteAndFormat 13
480                 ElseIf pPasteAs = "TEXT" Or pPasteAs = "TXT" Or pPasteAs = "T" Then
490                     wdDoc.Application.Selection.PasteAndFormat 22
500                 ElseIf pPasteAs = "TABLE" Or pPasteAs = "TBL" Then
                        'Esse metodo cola um conteudo copiado do excel, por exemplo, como formato de tabela
510                     wdDoc.Application.Selection.PasteExcelTable False, False, False    'Cola como tabela
520                 End If
530             End If
540         End If
550     End If
        Call AddAttachImages(p_oMail, Me.EmbeddedImages)
        Call EnviarMail(strAction)
        Set MailObject = p_oMail
660 End With

670 Enviar = Err = 0

Fim:
    If bCloseEmail Then
690     Set p_oMail = Nothing
    End If
700 If sTempFile <> "" Then If VBA.Dir(sTempFile) <> "" Then VBA.Kill sTempFile

710 Exit Function

Err_Outlook:
720 If VBA.Err <> 0 Then
        If VBA.Err.Number = 429 Then
            Debug.Print "Outlook esta fechado. Necessário que esteja aberto. Força a abertura do mesmo"
            Call StartOutlook
            GoTo CreateEmail:
        End If
730     Debug.Print " (" & VBA.Err.Description & ")"
740 End If
End Function

Private Function EnviarMail(pAction As eMailActions) As Boolean
1   On Error Resume Next
    m_oValidacaoDestinatarios = Me.ValidaRecipientes(p_oMail)
2   If pAction = Email_Send Then
3       p_oMail.send
4   ElseIf pAction = Email_Display Then
5       p_oMail.Display
6   ElseIf pAction = Email_SalvarRascunho Then
8       Call p_oMail.Close(0)
9   End If
10  EnviarMail = VBA.Err = 0
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cEmail.ValidaRecipientes()
' TIPO             : Function
' DATA/HORA        : 28/03/2014 20:21
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Verifica se os recipientes do email é valido
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      02/07/2016 17:01
'---------------------------------------------------------------------------------------
Function ValidaRecipientes(oMail As Object, Optional showTela As Boolean = False) As Variant
    Dim cOK As Collection, cInvalid As Collection, M
    Dim cErr As Collection
    Dim bResolved As Boolean
    Dim Rec As Object
    Dim fM As Object
    '---------------------------------------------------------------------------------------
1   On Error GoTo ValidaRecipientes_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
2   Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function cEmail.ValidaRecipientes()"
    '---------------------------------------------------------------------------------------
3   Set cOK = New Collection
4   Set cInvalid = New Collection
5   Set cErr = New Collection

6   For Each Rec In oMail.Recipients
7       bResolved = Rec.Resolve
8       If bResolved Then
9           cOK.Add Rec.Name & ";" & Rec.AddressEntry.Address & ";" & bResolved
10      Else
11          cErr.Add Rec
12          cInvalid.Add Rec.Name & ";" & Rec.Address & ";" & bResolved
13      End If
14  Next

15  If cInvalid.count > 0 Then
16      For Each Rec In cErr
17          'oMail.Recipients.Remove Rec.Index
18      Next
    End If

39  ValidaRecipientes = VBA.Array(cOK, cInvalid)

Fim:
40  On Error GoTo 0
41  Exit Function

ValidaRecipientes_Error:
42  If Err <> 0 Then
43      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
45  End If
    GoTo Fim:
    'Debug Mode
46  Resume
End Function
'--------------------------------------------------------------------------------------------------------------------------------
'Autor      : Adelson Silva - Mondial Informatica (adelsons@gmail.com)
'Data/Hora  : 02/09/2013 11:21
'Descrição  : Preeche algumas coleções do Outlook com varios Items
'             Algumas coleções interessantes são :
'              Recipients           : Permite adicionar varios endereços To, CC ou BCC
'              Attachments          : Permite adicionar varios endereços de arquivos que serão anexados
'---------------------------------------------------------------------------------------------------------------------------------
'02/09/2013 - Criado a função
'04/06/2013 - Parametrizada para que possa ser possivel preencher qual quer função com a mesma estrutura
'---------------------------------------------------------------------------------------------------------------------------------

Private Sub FillMailItems(strItems As String, col As Object, Optional itemType As Long)
    Dim item As String, TempArray, varArrayItem
    Dim oRecipient As Object
1   TempArray = VBA.Split(strItems, ";")
2   For Each varArrayItem In TempArray
3       item = VBA.Trim(varArrayItem)
4       If VBA.Len(item) > 0 Then
5           Set oRecipient = col.Add(item)
6           If itemType <> 0 Then oRecipient.Type = itemType
7           Set oRecipient = Nothing
8       End If
9   Next varArrayItem
End Sub

Function addInfo(ByRef vCampos, strNome As String, strValor As String)
    Dim vArr() As String
    If Not VBA.IsArray(vCampos) Then
        ReDim Preserve vArr(0 To 1, 0 To 0)
        vCampos = vArr
    Else
1       ReDim Preserve vCampos(0 To 1, 0 To UBound(vCampos, 2) + 1)
    End If
2   vCampos(0, UBound(vCampos, 2)) = strNome
3   vCampos(1, UBound(vCampos, 2)) = strValor
    Call VariaveisCorpo.Add(strNome & "|" & strValor)
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : BLPublica.PreencherVariaveis()
' TIPO             : Function
' DATA/HORA        : 31/05/2015 12:26
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Substituiu variáveis em um texto baseado em uma matriz de valores
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      31/05/2015 12:26    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Function PreencherMarcadores(cMarcadores As Collection) As String
    Dim iCtr As Integer
    Dim strValue As String
    Dim vl, strCampo As String
    Dim strMuliSelecao As String
    Dim wdDoc As Object
    Dim marcador As Object
    
10  On Error GoTo PreencherVariaveis_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function PreencherVariaveis()"
    '----------------------------------------------------------------------------------------------------
    If cMarcadores.count = 0 Then Exit Function
    p_oMail.Display
    Set wdDoc = p_oMail.GetInspector.WordEditor

50  For iCtr = 1 To cMarcadores.count
        strCampo = VBA.Split(cMarcadores(iCtr), "|")(0)
        strValue = VBA.Split(cMarcadores(iCtr), "|")(1)
        Set marcador = Nothing
        On Error Resume Next
        Set marcador = wdDoc.Bookmarks(strCampo)
        If Not marcador Is Nothing Then
            marcador.Range.Text = strValue
            'Debug.Print marcador.name, wdDoc.Range(marcador.Start, marcador.Start + Len(strValue))
            Do While wdDoc.Range(marcador.Start, marcador.Start + Len(strValue)) <> strValue
                VBA.DoEvents
                Debug.Print "Aguarda..."
                marcador.Range.Text = strValue
            Loop
            'Deleta o ultimo ] antes
            wdDoc.Range(marcador.Start + Len(strValue), marcador.Start + Len(strValue) + 1).Delete
            'Deleta o primeiro [
            wdDoc.Range(marcador.Start - 1, marcador.End).Delete
             ' On Error GoTo PreencherVariaveis_Error
        End If
270 Next iCtr

    p_oMail.Display
    Call p_oMail.Close(0)

'    p_oMail.Save
Fim:
290 On Error GoTo 0
300 Exit Function

PreencherVariaveis_Error:
310 If Err <> 0 Then
320     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
330     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
340 End If
    GoTo Fim:
    'Debug Mode
350 Resume

End Function


'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : BLPublica.PreencherVariaveis()
' TIPO             : Function
' DATA/HORA        : 31/05/2015 12:26
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Substituiu variáveis em um texto baseado em uma matriz de valores
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      31/05/2015 12:26    Adelson         Criação/Atualização do procedimento
'---------------------------------------------------------------------------------------
Function PreencherVariaveis(pTexto As String, cVariaveis As Collection) As String
    Dim iCtr As Integer
    Dim strValue As String
    Dim vl, strCampo As String
    Dim strMuliSelecao As String
    
10  On Error GoTo PreencherVariaveis_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function PreencherVariaveis()"
    '----------------------------------------------------------------------------------------------------
    If cVariaveis.count = 0 Then Exit Function

50  For iCtr = 1 To cVariaveis.count
90      strCampo = "[@" & VBA.Split(cVariaveis(iCtr), "|")(0) & "]"
110     strValue = VBA.Split(cVariaveis(iCtr), "|")(1)
120     If VBA.InStr(pTexto, strCampo) > 0 Then pTexto = VBA.Replace(pTexto, strCampo, strValue)
270 Next iCtr

280 PreencherVariaveis = pTexto

Fim:
290 On Error GoTo 0
300 Exit Function

PreencherVariaveis_Error:
310 If Err <> 0 Then
320     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
330     Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl)
340 End If
    GoTo Fim:
    'Debug Mode
350 Resume

End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cEmail.AddAttachImages()
' TIPO             : Function
' DATA/HORA        : 17/11/2016 15:17
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Adiciona imagems embarcadas no corpo do email
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      17/11/2016 15:17
'---------------------------------------------------------------------------------------
Function AddAttachImages(pMail As Object, oColImages As Collection)
    Dim wdDoc As Object    'Word.Document
    Dim strImagePath As String
    Dim intImg As Integer
    Dim vTags, sID As String, bHR As Boolean
    Dim docPosition As Object    'InlineShapes    ' Object
    Dim img As Object    ' InlineShape
    Dim sHTMLBody As String
    Dim M As Object ' MailItem
    Dim sBody As String


    '---------------------------------------------------------------------------------------
1   On Error GoTo AddAttachImages_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
2   Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Function cEmail.AddAttachImages()"
    '---------------------------------------------------------------------------------------
    If oColImages.count = 0 Then Exit Function
    
3   Set wdDoc = pMail.GetInspector.WordEditor
    pMail.Display
4   With pMail.GetInspector
5       .WindowState = 2
6       .Height = 0
7       .Width = 0
8   End With

9   For intImg = 1 To oColImages.count
10      GoSub GetInfoTagImage
11      If bHR Then
12          sBody = pMail.HTMLBody
13          If InStr(sBody, "<html>") = 0 Then
14              sHTMLBody = "<html><body>" & vbNewLine
15              sHTMLBody = sHTMLBody & sBody & vbNewLine
16              sHTMLBody = sHTMLBody & "<hr>" & vbNewLine
17              sHTMLBody = sHTMLBody & "</body>" & vbNewLine
18              sHTMLBody = sHTMLBody & "</html>" & vbNewLine
19          Else
20              sHTMLBody = sHTMLBody & sBody & vbNewLine
21              sHTMLBody = sHTMLBody & "<hr>" & vbNewLine
22          End If
23          M.HTMLBody = sHTMLBody
24      End If
25      Set docPosition = wdDoc.Content.Characters(wdDoc.Content.Characters.count).InlineShapes
27      Set img = docPosition.AddPicture(FileName:=strImagePath, LinkToFile:=False, SaveWithDocument:=True)
        'img.ConvertToShape
28      wdDoc.Content.Characters(wdDoc.Content.Characters.count).InsertBreak


29  Next intImg

30  Exit Function

GetInfoTagImage:
31  vTags = VBA.Split(oColImages(intImg), ";")
32  strImagePath = VBA.Split(vTags(0), "=")(1)
33  sID = VBA.Split(vTags(1), "=")(1)
34  bHR = VBA.Split(vTags(2), "=")(1)
35  Return

Fim:
36  On Error GoTo 0
37  Exit Function

AddAttachImages_Error:
38  If Err <> 0 Then
39      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
40      Debug.Print cstr_ProcedureName, "Linha : " & VBA.Erl() & " - " & strErrorMessagem
        'Call Excecoes.TratarErro(strErrorMessagem, lngErrorNumber, cstr_ProcedureName, VBA.Erl())
41  End If
    GoTo Fim:
    'Debug Mode
42  Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cEmail.getDefaultSignature()
' TIPO             : Function
' DATA/HORA        : 03/07/2016 14:26
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Recupera o contreduod HTML do arquiov de assinatura e insere no email gerado
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      03/07/2016 14:26
'---------------------------------------------------------------------------------------
Function getDefaultSignature() As String
    Dim sFileSign As String
    Dim sPathAppData As String
    sPathAppData = VBA.Environ("appdata") & "\Microsoft\Signatures\"
    sFileSign = VBA.Dir(sPathAppData & "*.htm")
    If sFileSign <> "" Then
        getDefaultSignature = ReadTextFile(sPathAppData & sFileSign)
    End If
End Function

'--------------------------------------------------------------------------------------------------------------------------------
'Autor      : Adelson Silva - Mondial Informatica (adelsons@gmail.com)
'Data/Hora  : 03/09/2013 12:10
'Descrição  : Abre um arquivo texto e obtem todo o seu conteudo em memoria
'---------------------------------------------------------------------------------------------------------------------------------
'03/09/2013 - Criado a função
'---------------------------------------------------------------------------------------------------------------------------------
Private Function ReadTextFile(pFile As String) As String
          Dim sContent As String, fso As Object, ts As Object
1         Set fso = VBA.CreateObject("Scripting.FileSystemObject")
2         Set ts = fso.GetFile(pFile).OpenAsTextStream(1, -2)
3         sContent = ts.ReadAll
4         ts.Close
5         ReadTextFile = sContent
End Function

Public Property Get Assunto() As String
    Assunto = m_sAssunto
End Property

Public Property Let Assunto(ByVal sAssunto As String)
    m_sAssunto = sAssunto
End Property

Public Property Get cc() As String
    cc = m_sCC
End Property

Public Property Let cc(ByVal sCC As String)
    m_sCC = sCC
End Property

Public Property Get cco() As String
    cco = m_sCCO
End Property

Public Property Let cco(ByVal sCCO As String)
    m_sCCO = sCCO
End Property

Public Property Get OnBehalfOf() As String
    OnBehalfOf = m_sOnBehalfOf
End Property

Public Property Let OnBehalfOf(ByVal sOnBehalfOf As String)
    m_sOnBehalfOf = sOnBehalfOf
End Property

Public Property Get Body() As Variant
    Body = m_vBody
End Property

Public Property Let Body(ByVal sBody As Variant)
    m_vBody = sBody
End Property

Public Property Get Para() As String
    Para = m_sPara
End Property

Public Property Let Para(ByVal sPara As String)
    m_sPara = sPara
End Property

Public Property Get MailObject() As Object
    Set MailObject = p_oMail
End Property

Public Property Set MailObject(objMailObject As Object)
    Set p_oMail = objMailObject
End Property

Public Property Get Session() As Object
    Set Session = m_objSession
End Property

Public Property Set Session(objSession As Object)
    Set m_objSession = objSession
End Property

Public Property Get VariaveisCorpo() As Collection
    If m_objVariaveisCorpo Is Nothing Then Set m_objVariaveisCorpo = New VBA.Collection
    Set VariaveisCorpo = m_objVariaveisCorpo
End Property

Public Property Set VariaveisCorpo(objVariaveisCorpo As Collection)
    Set m_objVariaveisCorpo = objVariaveisCorpo
End Property

Public Property Get VariaveisAssunto() As Collection
    If m_objVariaveisAssunto Is Nothing Then Set m_objVariaveisAssunto = New VBA.Collection
    Set VariaveisAssunto = m_objVariaveisAssunto
End Property

Public Property Set VariaveisAssunto(objVariaveisAssunto As Collection)
    Set m_objVariaveisAssunto = objVariaveisAssunto
End Property

Public Property Get MarcadoresCorpo() As Collection
    If m_objMarcadoresCorpo Is Nothing Then Set m_objMarcadoresCorpo = New VBA.Collection
    Set MarcadoresCorpo = m_objMarcadoresCorpo
End Property

Public Property Set MarcadoresCorpo(objMarcadoresCorpo As Collection)
    Set m_objMarcadoresCorpo = objMarcadoresCorpo
End Property

Public Property Get ValidacaoDestinatarios() As Variant
    ValidacaoDestinatarios = m_oValidacaoDestinatarios
End Property
