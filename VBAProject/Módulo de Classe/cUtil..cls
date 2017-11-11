'---------------------------------------------------------------------------------------
' PROCEDIMENTO     : Revisar_AuxFSO.CriarDiretorio()
' DATA/HORA        : 15/02/2016 15:51
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Cria toda a estrutura de pastas de um diretório informado
'---------------------------------------------------------------------------------------
'
' + Historico de Revisão
' **************************************************************************************
'   Versão    Data/Hora           Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      15/02/2016 15:51
' * 1.02      05/10/2016 10:20    Adelson          Redução, organização e otimização
' * 1.03      10/07/2017 10:10    Adelson          Identificado e corrigido problema na
'                                                  criação de pastas que contem o endereço
'                                                  UNC com o servidor começando em \\Servidor
'---------------------------------------------------------------------------------------
Function CriarDiretorio(ByVal strDir As String, Optional bSilent As Boolean = True) As Boolean
          '----------------------------------------------------------------------------------------------------
10        On Error GoTo TratarErro: Const cstr_ProcedureName As String = "AuxFSO.CriarDiretorio"
          '------------------------------------------------------------------------------------------------
          Dim sDiretorio As String
          Dim vDiretorios As Variant
          Dim i As Integer
          Dim sRoot As String
          Dim sDemais As String

20        If VBA.Left(strDir, 2) = "\\" Then
30            sRoot = "\\" & VBA.Split(VBA.Mid(strDir, 3), "\")(0)
40        Else
50            sRoot = VBA.Split(strDir, "\")(0)
60        End If

70        sDemais = VBA.Mid(strDir, Len(sRoot) + 1)
80        strDir = sDemais

90        vDiretorios = VBA.Split(strDir, "\")

100       On Error Resume Next
110       For i = LBound(vDiretorios) To UBound(vDiretorios)
              'If i = 0 Then sDiretorio = sRoot & "\" & vDiretorios(i)
120           sDiretorio = VBA.IIf(sDiretorio = "", vDiretorios(i), sDiretorio & "\" & vDiretorios(i))
              'Cria a pasta, caso ainda não exista
130           If VBA.Dir(sRoot & "\" & sDiretorio, VBA.vbDirectory) = "" Then
140               Call VBA.MkDir(sRoot & "\" & sDiretorio)
150           End If
160       Next

170       If VBA.Dir(sDiretorio, VBA.vbDirectory) = "" Then
180           If Not bSilent Then VBA.MsgBox "Criada com sucesso !", VBA.vbInformation
190           CriarDiretorio = True
200       Else
210           If Not bSilent Then VBA.MsgBox "Erro ao tentar cria a estrutura de pasta!" & VBA.Chr(10) & sDiretorio, VBA.vbCritical
220           CriarDiretorio = False
230       End If
240       On Error GoTo 0
Fim:
250       Exit Function
TratarErro:
260       If VBA.Err <> 0 Then Call VBA.Err.Raise(VBA.Rnd(), cstr_ProcedureName & " - " & VBA.Erl() & ">" & VBA.Err.Description & VBA.vbCrLf)
270       GoTo Fim:
280       Resume
End Function
