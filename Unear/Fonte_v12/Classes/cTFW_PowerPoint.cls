VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTFW_PowerPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' PROJETO/MODULO   : VBAProject.FWoPowerPoint
' TIPO             : Módulo de classe
' DATA/HORA        : 02/12/2016 16:56
' CONSULTOR        : Paulo Gallo
' DESCRIÇÃO        : Módulo de classe do PowerPoint
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão do Módulo
' **************************************************************************************
'   Versão    Data/Hora             Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      02/12/2016 16:56
'---------------------------------------------------------------------------------------

Option Explicit

Public Enum PpSaveAsFileType
    ppSaveAsBMP = 19
    ppSaveAsDefault = 11
    ppSaveAsEMF = 23
    ppSaveAsGIF = 16
    ppSaveAsHTML = 12
    ppSaveAsHTMLDual = 14
    ppSaveAsHTMLv3 = 13
    ppSaveAsJPG = 17
    ppSaveAsMetaFile = 15
    ppSaveAsOpenXMLAddin = 30
    ppSaveAsOpenXMLPresentation = 24
    ppSaveAsOpenXMLPresentationMacroEnabled = 25
    ppSaveAsOpenXMLShow = 28
    ppSaveAsOpenXMLShowMacroEnabled = 29
    ppSaveAsOpenXMLTemplate = 26
    ppSaveAsOpenXMLTemplateMacroEnabled = 27
    ppSaveAsOpenXMLTheme = 31
    ppSaveAsPDF = 32
    ppSaveAsPNG = 18
    ppSaveAsPresentation = 1
    ppSaveAsRTF = 6
    ppSaveAsShow = 7
    ppSaveAsTemplate = 5
    ppSaveAsTIF = 21
    ppSaveAsWebArchive = 20
    ppSaveAsXPS = 33
End Enum

Private m_objApplication As Object
Private m_objPresentation As Object
Private m_objSlides As Object
Private m_objShapes As Object

Private Sub Class_Initialize()
    Call NewPowerPoint
End Sub

Private Sub Class_Terminate()
    Set m_objApplication = Nothing
    Set m_objPresentation = Nothing
    Set m_objSlides = Nothing
    Set m_objShapes = Nothing
End Sub

Public Function NewPowerPoint() As Object
    Set m_objApplication = VBA.CreateObject("PowerPoint.Application")
    Set NewPowerPoint = m_objApplication
End Function

Friend Property Get Application() As Object
    Set Application = m_objApplication
End Property

Friend Property Set Application(objApplication As Object)
    Set m_objApplication = objApplication
End Property

Friend Property Get Presentation() As Object
    Set Presentation = m_objPresentation
End Property

Friend Property Set Presentation(objPresentations As Object)
    Set m_objPresentation = objPresentations
End Property

Friend Property Get Slides() As Object
    Set Slides = m_objSlides
End Property

Friend Property Set Slides(objSlides As Object)
    Set m_objSlides = objSlides
End Property

Friend Property Get Shapes() As Object
    Set Shapes = m_objShapes
End Property

Friend Property Set Shapes(ObjShapes As Object)
    Set m_objShapes = ObjShapes
End Property

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : oPowerPoint.OpenPresentation()
' TIPO             : Function
' DATA/HORA        : 02/12/2016 17:12
' CONSULTOR        : TECNUN - Paulo Gallo
' DESCRIÇÃO        : Rotina abre um apresentação.
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      02/12/2016 17:12
'---------------------------------------------------------------------------------------

Public Function OpenPresentation(Optional ByVal pFileName As String, _
                                 Optional ByVal ReadOnly As Boolean = True, _
                                 Optional ByVal Untitled As Boolean = False, _
                                 Optional ByVal WithWindow As Boolean = False) As Object
          
          '---------------------------------------------------------------------------------------
1         On Error GoTo OpenPresentation_Error
          Dim lngErrorNumber As Long, strErrorMessagem As String
2         Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
          Const cstr_ProcedureName As String = "Function oPowerPoint.OpenPresentation()"
          '---------------------------------------------------------------------------------------
          
3         If m_objApplication Is Nothing Then Call NewPowerPoint
4         If Not VBA.IsNull(pFileName) And Not pFileName = VBA.vbNullString Then
5             Set m_objPresentation = m_objApplication.Presentations.Open(pFileName, ReadOnly, Untitled, WithWindow)
6             Set m_objSlides = m_objPresentation.Slides
7             Set OpenPresentation = m_objPresentation
8         Else
9             Set m_objPresentation = m_objApplication.Presentations.Add(1)
10            Set m_objSlides = m_objPresentation.Slides
11            If Not Me.SlideExists(1) Then
12                Me.AddSlides (1)
13            End If
14            Set m_objSlides = m_objPresentation.Slides
15            Set OpenPresentation = m_objPresentation
16        End If
          
Fim:
17    On Error GoTo 0
18    Exit Function

OpenPresentation_Error:
19    If Err <> 0 Then
20         lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
21         Call Excecoes.TratarErro(lngErrorNumber, strErrorMessagem, cstr_ProcedureName)
22    End If
    GoTo Fim:
         'Debug Mode
23    Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : oPowerPoint.AddSlides()
' TIPO             : Function
' DATA/HORA        : 05/12/2016 10:46
' CONSULTOR        : TECNUN - Paulo Gallo
' DESCRIÇÃO        : Rotina adiciona um novo slide
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      05/12/2016 10:46
'---------------------------------------------------------------------------------------

Public Function AddSlides(ByVal index As Variant) As Object
'1         Me.Presentation.Slides.Add Index:=Me.CountSlides + 1, Layout:=ppLayoutCustom
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : oPowerPoint.SaveAs()
' TIPO             : Sub
' DATA/HORA        : 05/12/2016 15:49
' CONSULTOR        : TECNUN - Paulo Gallo
' DESCRIÇÃO        : Rotina salva a apresentação.
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      05/12/2016 15:49
'---------------------------------------------------------------------------------------
Public Sub SaveAsPresentation(ByVal FileName, ByVal ext As PpSaveAsFileType)

          '---------------------------------------------------------------------------------------
1         On Error GoTo SaveAsPresentation_Error
          Dim lngErrorNumber As Long, strErrorMessagem As String
2         Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
          Const cstr_ProcedureName As String = "Sub oPowerPoint.SaveAsPresentation()"
          '---------------------------------------------------------------------------------------

3         If m_objApplication Is Nothing Then Call NewPowerPoint
4         If VBA.IsNull(FileName) And FileName = VBA.vbNullString Then
5              Call Me.Presentation.SaveAs(FileName & "\" & "Apresentação01", ext)
6         Else
7             Call Me.Presentation.SaveAs(FileName & "\" & "Apresentação01", ext, ext)
8         End If

Fim:
9     On Error GoTo 0
10    Exit Sub

SaveAsPresentation_Error:
11    If Err <> 0 Then
12         lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
13         Call Excecoes.TratarErro(lngErrorNumber, strErrorMessagem, cstr_ProcedureName)
14    End If
    GoTo Fim:
         'Debug Mode
15    Resume
          
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : oPowerPoint.CopySlide()
' TIPO             : Function
' DATA/HORA        : 05/12/2016 14:53
' CONSULTOR        : TECNUN - Paulo Gallo
' DESCRIÇÃO        : Rotina duplica slide informado.
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      05/12/2016 14:53
'---------------------------------------------------------------------------------------

Public Function DuplicateSlide(Optional ByVal index As Variant) As Object
1         If Not Me.SlideExists(index) Then Exit Function
2         Set DuplicateSlide = Me.Slides(index).Duplicate
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : oPowerPoint.DeleteSlide()
' TIPO             : Sub
' DATA/HORA        : 05/12/2016 16:23
' CONSULTOR        : TECNUN - Paulo Gallo
' DESCRIÇÃO        : Rotina exclui slide da apresentação.
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      05/12/2016 16:23
'---------------------------------------------------------------------------------------

Public Sub DeleteSlide(ByVal index As Variant)
1         If Not Me.SlideExists(index) Then Exit Sub
2         Me.Slides(index).Delete
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : oPowerPoint.QtdSlides()
' TIPO             : Function
' DATA/HORA        : 09/12/2016 11:48
' CONSULTOR        : TECNUN - Paulo Gallo
' DESCRIÇÃO        : Rotina retorna quantidade de slides da apresentação instanciada.
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      09/12/2016 11:48
'---------------------------------------------------------------------------------------

Public Function CountSlides() As Long
1         CountSlides = Me.Slides.count
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : oPowerPoint.ShowVisible()
' TIPO             : Sub
' DATA/HORA        : 09/12/2016 12:15
' CONSULTOR        : TECNUN - Paulo Gallo
' DESCRIÇÃO        : Rotina exibe objeto oculto
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      09/12/2016 12:15
'---------------------------------------------------------------------------------------

Public Sub ShowHidden(Optional ByVal bShow As Boolean = True)
1         Me.Application.Visible = bShow
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : oPowerPoint.SlideExists()
' TIPO             : Function
' DATA/HORA        : 09/12/2016 15:33
' CONSULTOR        : TECNUN - Paulo Gallo
' DESCRIÇÃO        : Rotina verifica se um slide existe na apresentação.
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      09/12/2016 15:33
'---------------------------------------------------------------------------------------

Public Function SlideExists(ByVal index As Variant) As Boolean
'
          '---------------------------------------------------------------------------------------
1         On Error GoTo SlideExists_Error
          Dim lngErrorNumber As Long, strErrorMessagem As String
2         Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
          Const cstr_ProcedureName As String = "Function oPowerPoint.SlideExists()"
          '---------------------------------------------------------------------------------------
          Dim sld     As Object
          
3         SlideExists = False
4         For Each sld In Me.Slides
5             If sld.slideNumber = index Or sld.Name = index Or sld.SlideIndex = index Then
6                 SlideExists = True
7                 GoTo Fim
8             End If
9         Next
Fim:
10    On Error GoTo 0
11    Call Publicas.RemoverObjetosMemoria(sld)
12    Exit Function

SlideExists_Error:
13    If Err <> 0 Then
14         lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
15         Call Excecoes.TratarErro(lngErrorNumber, strErrorMessagem, cstr_ProcedureName)
16    End If
    GoTo Fim:
         'Debug Mode
17    Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : oPowerPoint.CutPicture()
' TIPO             : Sub
' DATA/HORA        : 09/12/2016 15:38
' CONSULTOR        : TECNUN - Paulo Gallo
' DESCRIÇÃO        : Rotina corta uma imagem atraves dos parametros de tamanho da imagem e de corte.
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      09/12/2016 15:38
'---------------------------------------------------------------------------------------

Public Sub CutPicture(ByVal index As Variant, _
                           ByVal IndexShape As Variant, _
                           ByVal PictureWidth As Double, _
                           ByVal PictureHeight As Double, _
                           ByVal ShapeWidth As Double, _
                           ByVal ShapeHeight As Double, _
                           ByVal ShapeLeft As Double, _
                           ByVal ShapeTop As Double, _
                           ByVal PictureOffsetX As Double, _
                           ByVal PictureOffsetY As Double, _
                           ByVal Diametro As Double)

          '---------------------------------------------------------------------------------------
1         On Error GoTo CutPicture_Error
          Dim lngErrorNumber As Long, strErrorMessagem As String
2         Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
          Const cstr_ProcedureName As String = "Sub oPowerPoint.CutPicture()"
          '---------------------------------------------------------------------------------------

3         With Me.Slides(index).Shapes(IndexShape).PictureFormat.Crop
4             .PictureHeight = PictureHeight * Diametro
5             .PictureWidth = PictureWidth * Diametro
6             .ShapeHeight = ShapeHeight * Diametro
7             .ShapeWidth = ShapeWidth * Diametro
8             .ShapeLeft = ShapeLeft * Diametro
9             .ShapeTop = ShapeTop * Diametro
10            .PictureOffsetX = PictureOffsetX * Diametro
11            .PictureOffsetY = PictureOffsetY * Diametro
12        End With
          
Fim:
13    On Error GoTo 0
14    Exit Sub

CutPicture_Error:
15    If Err <> 0 Then
16         lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
17         Call Excecoes.TratarErro(lngErrorNumber, strErrorMessagem, cstr_ProcedureName)
18    End If
    GoTo Fim:
         'Debug Mode
19    Resume
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : oPowerPoint.AddPicture()
' TIPO             : Function
' DATA/HORA        : 12/12/2016 09:27
' CONSULTOR        : TECNUN - Paulo Gallo
' DESCRIÇÃO        : Rotina Adiciona um imagem a um apresentação.
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      12/12/2016 09:27
'---------------------------------------------------------------------------------------

Public Function AddPicture(ByVal pFileName As String, _
                           ByVal index As Variant, _
                           ByVal ExtensaoFile As String, _
                           ByVal LinkToFile As Boolean, _
                           ByVal SaveWithDocument As Boolean, _
                           ByVal Diametro As Double, _
                           Optional ByVal Left As Double = 0, _
                           Optional ByVal Top As Double = 0, _
                           Optional ByVal Width As Double = 0, _
                           Optional ByVal Height As Double = 0) As Object

          '---------------------------------------------------------------------------------------
1         On Error GoTo AddPicture_Error
          Dim lngErrorNumber As Long, strErrorMessagem As String
2         Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
          Const cstr_ProcedureName As String = "Function oPowerPoint.AddPicture()"
          '---------------------------------------------------------------------------------------
          
3         If Not pFileName = VBA.vbNullString And Not VBA.IsEmpty(pFileName) Then
4         Select Case ExtensaoFile
              Case ".png", ".jpg"
5                 Set AddPicture = Me.Slides(index).Shapes.AddPicture(pFileName & ExtensaoFile, False, True, Left * Diametro, _
                                             Top * Diametro, Width * Diametro, Height * Diametro)
6         End Select
7         End If
          
Fim:
8     On Error GoTo 0
9     Exit Function

AddPicture_Error:
10    If Err <> 0 Then
11         lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
12         Call Excecoes.TratarErro(lngErrorNumber, strErrorMessagem, cstr_ProcedureName)
13    End If
    GoTo Fim:
         'Debug Mode
14    Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cTFW_PowerPoint.GetRedimPicture()
' TIPO             : Function
' DATA/HORA        : 19/01/2017 14:16
' CONSULTOR        : TECNUN - Paulo Gallo
' DESCRIÇÃO        : Rotina pega a dimensão do objeto e salva em um array
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      19/01/2017 14:16
'---------------------------------------------------------------------------------------
Public Function GetRedimPicture(ByVal ObjShapes As String, _
                                ByVal index As Integer, _
                                Optional DeleteShape As Boolean = True) As Variant

          '---------------------------------------------------------------------------------------
1         On Error GoTo GetRedimPicture_Error
          Dim lngErrorNumber As Long, strErrorMessagem As String
2         Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
          Const cstr_ProcedureName As String = "Function cTFW_PowerPoint.GetRedimPicture()"
          '---------------------------------------------------------------------------------------

          Dim arrRedimPicture         As Variant

3         ReDim arrRedimPicture(0 To 3)
4         With Me.Presentation.Slides(index).Shapes(ObjShapes)
5             arrRedimPicture(0) = .Left
6             arrRedimPicture(1) = .Top
7             arrRedimPicture(2) = .Height
8             arrRedimPicture(3) = .Width
9         End With
          
10        If DeleteShape Then Me.Presentation.Slides(index).Shapes(ObjShapes).Delete
11        GetRedimPicture = arrRedimPicture
          
Fim:
12    On Error GoTo 0
13    Exit Function

GetRedimPicture_Error:
14    If Err <> 0 Then
15         lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
16         Call Excecoes.TratarErro(lngErrorNumber, strErrorMessagem, cstr_ProcedureName)
17    End If
    GoTo Fim:
         'Debug Mode
18    Resume
End Function

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cTFW_PowerPoint.SetRedimPicture()
' TIPO             : Sub
' DATA/HORA        : 19/01/2017 17:02
' CONSULTOR        : TECNUN - Paulo Gallo
' DESCRIÇÃO        : Rotina atribuiu a um shape dimensões
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      19/01/2017 17:02
'---------------------------------------------------------------------------------------

Public Sub SetRedimPicture(ByVal ObjShapes As String, ByVal index As Integer, ByRef arrSetRedimPicture As Variant)
          
          
          '---------------------------------------------------------------------------------------
1         On Error GoTo SetRedimPicture_Error
          Dim lngErrorNumber As Long, strErrorMessagem As String
2         Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
          Const cstr_ProcedureName As String = "Sub cTFW_PowerPoint.SetRedimPicture()"
          '---------------------------------------------------------------------------------------

3         If Not VBA.IsArray(arrSetRedimPicture) Then GoTo Fim
4         With Me.Presentation.Slides(index).Shapes(ObjShapes)
5             .LockAspectRatio = False
6             .Left = arrSetRedimPicture(0)
7             .Top = arrSetRedimPicture(1)
8             .Height = arrSetRedimPicture(2)
9             .Width = arrSetRedimPicture(3)
10        End With

Fim:
11    On Error GoTo 0
12    Exit Sub

SetRedimPicture_Error:
13    If Err <> 0 Then
14         lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
15         Call Excecoes.TratarErro(lngErrorNumber, strErrorMessagem, cstr_ProcedureName)
16    End If
    GoTo Fim:
         'Debug Mode
17    Resume
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cTFW_PowerPoint.ClosePresentation()
' TIPO             : Sub
' DATA/HORA        : 20/01/2017 01:21
' CONSULTOR        : Paulo Gallo
' DESCRIÇÃO        : Rotina fecha Apresentacao ativa.
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      20/01/2017 01:21
'---------------------------------------------------------------------------------------
Public Sub ClosePresentation(ByVal objPPT As Object, _
                             Optional ByVal bSalve As Boolean = True)

    '---------------------------------------------------------------------------------------
    On Error GoTo ClosePresentation_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
    Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Sub cTFW_PowerPoint.ClosePresentation()"
    '---------------------------------------------------------------------------------------
    
    If objPPT Is Nothing Then
        With objPPT
            If bSalve Then .Save
            .Close
        End With
    End If
Fim:
On Error GoTo 0
Exit Sub

ClosePresentation_Error:
If Err <> 0 Then
     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
     Call Excecoes.TratarErro(lngErrorNumber, strErrorMessagem, cstr_ProcedureName)
End If
GoTo Fim:
'Debug Mode
Resume
End Sub

'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : cTFW_PowerPoint.Shapes_PasteSpecial()
' TIPO             : Sub
' DATA/HORA        : 30/05/2017 08:54
' CONSULTOR        : Paulo Gallo
' DESCRIÇÃO        : Rotina verifica qual a melhor opção para colar dados no Power Point verificando o Type do Shape.
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      20/01/2017 01:21
'---------------------------------------------------------------------------------------
Public Function Shapes_PasteSpecial(ByVal index As Variant, _
                                    ByVal ShapeType As Variant) As Object

    '---------------------------------------------------------------------------------------
    On Error GoTo Shapes_PasteSpecial_Error
    Dim lngErrorNumber As Long, strErrorMessagem As String
    Dim dtSartRunProc As Date: dtSartRunProc = VBA.Time
    Const cstr_ProcedureName As String = "Sub cTFW_PowerPoint.Shapes_PasteSpecial()"
    '---------------------------------------------------------------------------------------
    
    Select Case ShapeType
        Case 17, 7, 3, 6
            'Const ppPasteShape = 11
            Set m_objShapes = Me.Presentation.Slides(index).Shapes.PasteSpecial(DataType:=11).item(1)
        Case 19, 1
            'ppPasteEnhancedMetafile = 2
            Set m_objShapes = Me.Presentation.Slides(index).Shapes.PasteSpecial(DataType:=2).item(1)
    End Select
    Set Shapes_PasteSpecial = m_objShapes
Fim:
On Error GoTo 0
Exit Function

Shapes_PasteSpecial_Error:
If Err <> 0 Then
     lngErrorNumber = VBA.Err.Number: strErrorMessagem = VBA.Err.Description
     Call TratarErro(lngErrorNumber, strErrorMessagem, cstr_ProcedureName)
End If
GoTo Fim:
'Debug Mode
Resume
End Function

