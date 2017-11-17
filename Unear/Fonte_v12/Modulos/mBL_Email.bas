Attribute VB_Name = "mBL_Email"
Sub EnviarEmail(pmodelo As Integer, strAnexos As String)
    Dim email As cTFW_Email
    Dim strArquivoModelo As String
    Dim strNomeModelo As String
    Dim vCheckDestinatarios As Variant
    Dim lngDestinatario As Long
    Dim strOnBehalf As String
    Dim strEmailDestino As String
    
    Set cOutlookApp = New cTFW_Outlook
    
    Set email = New cTFW_Email
    strNomeModelo = Nz(DLookup("NomeModelo", "tblModelosEmails", "IDModelo = " & pmodelo), "ModeloNaoEncontrado")
    strArquivoModelo = PegaEndereco_Templates & "\" & strNomeModelo & ".msg"
    strOnBehalf = CaixaDeCorreioRemetente()
    strEmailDestino = EmailsLogs()
    
    Call Inicializar_Globais
    
    With email
        .Para = pegaValor("EmailDestino")
        'Prepara o email
        Call .CriarNovoEmail(strArquivoModelo)
        'Carrega as variáveis do modelo
        Set rsLista = Conexao.PegarRS("pegarVariaveisModelo", pmodelo)
        Do While Not rsLista.EOF
            Call .MarcadoresCorpo.Add(Nz(rsLista!CampoVariavel.value) & "|" & Nz(rsLista!ValorVariavel.value), Nz(rsLista!CampoVariavel.value))
            Call .VariaveisAssunto.Add(Nz(rsLista!CampoVariavel.value) & "|" & Nz(rsLista!ValorVariavel.value))
            rsLista.MoveNext
        Loop
        If strOnBehalf <> "" Then .OnBehalfOf = strOnBehalf
        .Para = strEmailDestino
        Call showMessage(sMsg:="Enviando email para : " & VBA.CStr(vEmail), modoProgress:=EmEndamento_ProgressoFixo)
        '### DISPARA O ENVIO ###
        Call .Enviar(strAction:=eMailActions.Email_Send, bCloseEmail:=True, strAttachmentPaths:=strAnexos)
        vCheckDestinatarios = .ValidacaoDestinatarios
    End With
    Call showMessage(sMsg:="Envio dos emails concluído !", modoProgress:=Fim, intTimer:=2)
End Sub

Sub AtualizaModelo(pIDModelo, campo As String, valor As String)
    With CurrentDb.OpenRecordset("SELECT * FROM tblModelosEmails WHERE IDModelo = " & pIDModelo)
        If .EOF Then .addNew Else .edit
        .Fields(campo).value = valor
        .Update
    End With
End Sub

Sub AtualizaVariavelModelo(pIDModelo, IDVariavel As String, valor As String)
    With CurrentDb.OpenRecordset("SELECT * FROM tblVariaveisModelos WHERE IDModelo = " & pIDModelo & " and CampoVariavel = '" & IDVariavel & "'")
        If .EOF Then .addNew Else .edit
        .Fields("IDModelo").value = pIDModelo
        .Fields("CampoVariavel").value = IDVariavel
        .Fields("ValorVariavel").value = valor
        .Update
    End With
End Sub


