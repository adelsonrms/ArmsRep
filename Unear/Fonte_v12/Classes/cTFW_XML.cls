VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTFW_XML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Private Const XML_HEADER_UTF_8 As String = "<?xml version=""1.0"" encoding=""utf-8""?>"
Private Const XML_HEADER_ISO As String = "<?xml version=""1.0"" encoding=""ISO-8859-1""?>"

Public Enum FormatoXML
    XML_PORTUGUES = 1
    XML_GLOBAL = 2
End Enum

Private xDoc As Object


Function openDocument(pSource As String)
    Set openDocument = GetXDoc(pSource)
End Function

Private Function GetXDoc2(Optional pXMLSource As Variant) As Object
    Set xDoc = VBA.CreateObject("MSXML2.DOMDocument")
    If Not VBA.IsMissing(pXMLSource) Then
        If VBA.varType(pXMLSource) = vbObject Then
            Set xDoc = pXMLSource
        Else
            If Left(Trim(pXMLSource), 1) = "<" And Right(Trim(pXMLSource), 1) = ">" Then
                'Definição XML
                Call xDoc.LoadXML
            Else
                'Arquivo ou URL
                Call xDoc.Load(pXMLSource)
            End If
        End If
    End If
    Set GetXDoc2 = xDoc
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : mdlXML.GetXDoc()
' TIPO             : Function
' DATA/HORA        : 05/11/2014 14:27
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Criar ou abre um documento XML. A partir de um arquivo existente ou de um conteudo XML
'---------------------------------------------------------------------------------------
'REVISÕES
'---------------------------------------------------------------------------------------
'10/11/2014 14:19   - Instanciar com Late Bounding. Usando MSXML.DOMDocument
Function GetXDoc(pSource As String, Optional bCreateNew As Boolean = False, Optional pRootElement As String = "") As Object
    '----------------------------------------------------------------------------------------------------
10  On Error GoTo GetXDoc_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = Time
    Const cstr_ProcedureName As String = "Function prjArmsFinanceManager.mdlXML.GetXDoc()"
    '----------------------------------------------------------------------------------------------------
30  Set xDoc = VBA.CreateObject("MSXML.DOMDocument")

    If pRootElement = "" Then pRootElement = "ROOT"

40  If InStr(pSource, "<") > 0 And InStr(pSource, "/>") > 0 Then
50      Call xDoc.LoadXML(pSource)
    ElseIf pSource = "" Then
        Call xDoc.appendChild(xDoc.createElement(pRootElement))
60  ElseIf VBA.Dir(pSource) <> "" Then
70      If bCreateNew Then Call VBA.Kill(pSource)
80      Call xDoc.Load(pSource)
90      If pRootElement <> "" Then
100         If xDoc.DocumentElement Is Nothing Then Call xDoc.appendChild(xDoc.createElement(pRootElement))
110     End If
120 End If
    
    Set GetXDoc = xDoc

130 On Error GoTo 0
140 Exit Function

GetXDoc_Error:
150 If Err <> 0 Then
160     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        'Quando não há o modulo de exceções
170     Call VBA.MsgBox("Ocorreu um erro ao criar a nova tarefa" & vbNewLine & vbNewLine & _
                        "Detalhes : " & cstr_ProcedureName & "." & VBA.Erl & " - " & lngErrorNumber & "-" & strErrorMessagem)
180 End If
190 Exit Function
    'Debug Mode
200 Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : mdlXML.addElement()
' TIPO             : Function
' DATA/HORA        : 05/11/2014 14:12
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Adiciona um elemento no XML
'---------------------------------------------------------------------------------------
Function AddElement(xParent As Object, pName As String, Optional pValue As String, Optional pSave As Boolean = False)
    Dim xInfo  As Object
    '----------------------------------------------------------------------------------------------------
10  On Error GoTo addElement_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = Time
    Const cstr_ProcedureName As String = "Function prjArmsFinanceManager.mdlXML.addElement()"
    '----------------------------------------------------------------------------------------------------

30  Set xInfo = xParent.ownerDocument.createElement(pName)
40  xInfo.Text = pValue
50  Call xParent.appendChild(xInfo)

60  On Error GoTo 0
70  Exit Function

addElement_Error:
80  If Err <> 0 Then
90      lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        'Quando não há o modulo de exceções
100     Call VBA.MsgBox("Ocorreu um erro ao criar a nova tarefa" & vbNewLine & vbNewLine & _
                        "Detalhes : " & cstr_ProcedureName & "." & VBA.Erl & " - " & lngErrorNumber & "-" & strErrorMessagem)
110 End If
120 Exit Function
    'Debug Mode
130 Resume
End Function

Function getElement(pXPath As String, Optional objParent As DOMDocument) As Object
    Dim xElement As Object
    If objParent Is Nothing Then Set objParent = xDoc
    Set xElement = objParent.selectNodes(VBA.Replace(pXPath, "\", "/"))
    Set getElement = xElement.item(0)
End Function

Function getElements(pXPath As String, Optional objParent As DOMDocument) As Object
    Dim xElement As IXMLDOMNodeList
    If objParent Is Nothing Then Set objParent = xDoc
    Set xElement = objParent.selectNodes(pXPath)
    Set getElements = xElement
End Function

Function getValue(pXPath As String, Optional objParent As Object)
    Dim nd As Object, xTextValue
    If objParent Is Nothing Then
        Set objParent = xDoc
    End If
    
9   Set nd = objParent.selectNodes(pXPath)
10  If Not nd Is Nothing Then
        If nd.Length = 0 Then GoTo Finaliza
11      If nd.Length > 1 Then
            If vRetorno = vbObject Then
12              Set xTextValue = nd
            Else
                xTextValue = nd.item(0).Text
            End If
13      Else
14          xTextValue = nd.item(0).Text
15      End If
16  End If
Finaliza:
24      getValue = xTextValue
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : mdlXML.FormatXML()
' TIPO             : Function
' DATA/HORA        : 05/11/2014 14:13
' CONSULTOR        : (ADELSON)/ TECNUN - Adelson Rosendo Marques da Silva
' CONTATO          : adelson@tecnun.com.br
' DESCRIÇÃO        : Formata um conteúdo XML
'---------------------------------------------------------------------------------------
Public Function FormatXML(XML As String) As String
    '----------------------------------------------------------------------------------------------------
10  On Error GoTo FormatXML_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
20  Dim dtSartRunProc As Date: dtSartRunProc = Time
    Const cstr_ProcedureName As String = "Function prjArmsFinanceManager.mdlXML.FormatXML()"
    '----------------------------------------------------------------------------------------------------

    Dim Reader As Object    'SAXXMLReader60
    Dim Writer As Object    'MXXMLWriter60

30  Set Reader = VBA.CreateObject("MSXML2.SAXXMLReader60")
40  Set Writer = VBA.CreateObject("MSXML2.MXXMLWriter60")

50  With Writer
60      .indent = True
70      .standalone = False
80      .omitXMLDeclaration = False
90      .Encoding = "utf-8"
100 End With

110 With Reader
120     Set .contentHandler = Writer
130     Set .dtdHandler = Writer
140     Set .ErrorHandler = Writer
150     Call .putProperty("http://xml.org/sax/properties/declaration-handler", Writer)
160     Call .putProperty("http://xml.org/sax/properties/lexical-handler", Writer)
170     Call .Parse(XML)
180 End With

190 FormatXML = Writer.Output

200 On Error GoTo 0
210 Exit Function

FormatXML_Error:
220 If Err <> 0 Then
230     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
        'Quando não há o modulo de exceções
240     Call VBA.MsgBox("Ocorreu um erro ao criar a nova tarefa" & vbNewLine & vbNewLine & _
                        "Detalhes : " & cstr_ProcedureName & "." & VBA.Erl & " - " & lngErrorNumber & "-" & strErrorMessagem)
250 End If
260 Exit Function
    'Debug Mode
270 Resume
End Function

Function setSchema(pSchemaFile As String) As Object 'XMLSchemaCache60
    Dim oSC    As Object ' XMLSchemaCache60
    'Create the schema cache and add the XSD schema to it.
10  Set oSC = VBA.CreateObject("MSXML2.XMLSchemaCache.6.0")
20  oSC.Add "urn:Extract", pSchemaFile
30  Set setSchema = oSC
    'Set properties, load and validate it in the XML DOM.
End Function





