Attribute VB_Name = "M�dulo7"
Sub aaaaaaaaaaaa()

    Dim caminhoTemplate As String
    Dim dadosEstrutura As Object
    Dim corpoEmail As String
    Dim emailAssessor As String
    Dim emailCliente As String
    Dim assuntoEmail As String
    Dim nomeEstrutura As String
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Teste")
    
    nomeEstrutura = ws.Range("G11").Value

    caminhoTemplate = IdentificarTemplate(nomeEstrutura)
    If caminhoTemplate = "" Then
        MsgBox "Template n�o encontrado para a estrutura: " & nomeEstrutura
        Exit Sub
    End If

    Set dadosEstrutura = CriarDict(nomeEstrutura, ws)

    corpoEmail = FormatarEmail(caminhoTemplate, dadosEstrutura)
    
    emailAssessor = ws.Range("E10").Value
    emailCliente = ws.Range("C10").Value

    assuntoEmail = "Opera��o " & nomeEstrutura & " - " & ws.Range("G11").Value

    montarEmail emailAssessor, emailCliente, corpoEmail, assuntoEmail
End Sub


Function FormatarEmail(caminhoArquivo As String, dados As Object) As String
    'Funcao para abrir e formatar aruivo de template de email
    'Args:
        '- caminhoArquivo: caminho do arquivo como path para ser aberto e formatado
        '- dados: dicionario contendo as informacoes para formatar o arquivo
    'Raises:
        '- 13: Type Mismatch, o argumento n�o � um dicion�rio
        '- 75: Path/File access error, n�o foi poss�vel l�r o arquivo
    'Returns:
        '- Texto do arquivo aberto e formatado

    'Declara argumento para key do dict
    Dim chave As Variant
    Dim numeroArquivo As Integer
    Dim conteudoArquivo As String
    
    'Checa se o argumento Dados � Dicion�rio
    If Not TypeName(dados) = "Dictionary" Then
        Err.Raise Number:=13, Description:="Type Mismatch: O argumento n�o � um dicion�rio v�lido."
    End If

    'Informa qual espaco de mem�ria usar
    numeroArquivo = FreeFile
    
    'Abrir o arquivo para leitura
    On Error GoTo ErroAbrirArquivo
    Open caminhoArquivo For Input As numeroArquivo
        'LOF l� o arquivo em bytes e aloca na vari�vel
        conteudoArquivo = Input(LOF(numeroArquivo), numeroArquivo)
    Close numeroArquivo

    'Formatar o arquivo
    For Each chave In dados.Keys
        ' Substituir cada placeholder pelo valor correspondente
        conteudoArquivo = Replace(conteudoArquivo, "{{" & chave & "}}", dados(chave))
    Next chave

    'Return da fun��o
    FormatarEmail = conteudoArquivo
    Exit Function

'Tratamento de erro ao ler arquivo
ErroAbrirArquivo:
    Err.Raise Number:=75, Description:="Path/File access error: N�o foi poss�vel abrir o arquivo"
End Function
Function IdentificarTemplate(valor As String) As String
    ' Fun��o para indentificar o template referente ao tipo de estrutura
    ' Args:
        '- Valor: tipo de estrutura escolhido pelo broker
        
    ' Returns:
        '- O caminho do arquivo HTML

    Dim caminho As String
    caminho = Environ("USERPROFILE") & "XP Investimentos\Manchester - Mesa RV - Backoffice - Backoffice\Projetos\Envio de email estruturadas boletera\tamplates\"

    Select Case valor
        Case "Aloca��o Protegida"
            IdentificarTemplate = caminho & "alocacaoprotegida.html"
        
        Case "Booster"
            IdentificarTemplate = caminho & "booster.html"
            
        Case "Booster Shield"
            IdentificarTemplate = caminho & "boostershield.html"
            
        Case "Collar UI"
            IdentificarTemplate = caminho & "collarui.html"
            
        Case "Financiamento"
            IdentificarTemplate = caminho & "financiamento.html"
            
        Case "NDF"
            IdentificarTemplate = caminho & "ndf.html"
            
        Case "NDF com CAP"
            IdentificarTemplate = caminho & "ndfcomcap.html"
            
        Case "Rubi"
            IdentificarTemplate = caminho & "rubi.html"
        
        Case Else
            IdentificarTemplate = ""
    End Select
End Function
Function CriarDict(nomeEstrutura As String, ws As Worksheet) As Object
    'Funcao para criar dicionario com dados de uma estrutura com
    'base no nome da estrutura
    
    'Args:
        '- nomeEstrutura: nome da na c�lula D11 usada para criar o dicion�rio
        '- ws: nome do sheets do arquivo Excel
    'Returns:
        '- dicion�rio com os argumentos dos arquivos HTML de cada estrutura

    'Declarando o dicionario
    
    Dim dados As Object
    Set dados = CreateObject("Scripting.Dictionary")

    Select Case nomeEstrutura
        Case "Aloca��o Protegida"
            dados.Add "ATIVO", ws.Range("J10").Value
            dados.Add "QUANTIDADE", ws.Range("K10").Value
            dados.Add "STRIKE", ws.Range("L10").Value
            dados.Add "PR�MIO", ws.Range("M10").Value
            dados.Add "PRE�O", ws.Range("N10").Value
            dados.Add "VENCIMENTO", ws.Range("O10").Value
            dados.Add "OPERA��O", ws.Range("P10").Value
        
        Case "Booster"
            dados.Add "ATIVO", ws.Range("J10").Value
            dados.Add "QUANTIDADE", ws.Range("K10").Value
            dados.Add "PRE�O REF", ws.Range("L10").Value
            dados.Add "VENCIMENTO", ws.Range("M10").Value
            dados.Add "STRIKE CALL VENDIDA", ws.Range("N10").Value
            dados.Add "STRIKE CALL COMPRADA", ws.Range("O10").Value
            dados.Add "OPERA��O", ws.Range("P10").Value

        Case "Booster Shield"
            dados.Add "ATIVO", ws.Range("J10").Value
            dados.Add "QUANTIDADE", ws.Range("K10").Value
            dados.Add "PRE�O REF", ws.Range("L10").Value
            dados.Add "VENCIMENTO", ws.Range("M10").Value
            dados.Add "STRIKE PUT COMPRADA", ws.Range("N10").Value
            dados.Add "STRIKE CALL VENDIDA", ws.Range("O10").Value
            dados.Add "STRIKE CALL COMPRADA", ws.Range("P10").Value
            dados.Add "BARREIRA", ws.Range("Q10").Value
            dados.Add "OPERA��O", ws.Range("R10").Value

        Case "Collar UI"
            dados.Add "ATIVO", ws.Range("J10").Value
            dados.Add "QUANTIDADE", ws.Range("K10").Value
            dados.Add "PRE�O", ws.Range("L10").Value
            dados.Add "VENCIMENTO", ws.Range("M10").Value
            dados.Add "STRIKE PUT", ws.Range("N10").Value
            dados.Add "STRIKE CALL", ws.Range("O10").Value
            dados.Add "BARREIRA", ws.Range("P10").Value
            dados.Add "OPERA��O", ws.Range("Q10").Value

        Case "Financiamento"
            dados.Add "ATIVO", ws.Range("J10").Value
            dados.Add "QUANTIDADE", ws.Range("K10").Value
            dados.Add "PRE�O", ws.Range("L10").Value
            dados.Add "VENCIMENTO", ws.Range("M10").Value
            dados.Add "STRIKE", ws.Range("N10").Value
            dados.Add "PR�MIO", ws.Range("O10").Value
            dados.Add "OPERA��O", ws.Range("P10").Value

        Case "NDF"
            dados.Add "PRE�O COMPRA", ws.Range("J10").Value
            dados.Add "PRE�O REF", ws.Range("K10").Value
            dados.Add "VENCIMENTO", ws.Range("L10").Value
            dados.Add "VOLUME", ws.Range("M10").Value
            dados.Add "DATA", ws.Range("N10").Value
            dados.Add "OPERA��O", ws.Range("O10").Value

        Case "NDF com CAP"
            dados.Add "PRE�O COMPRA", ws.Range("J10").Value
            dados.Add "PRE�O REF", ws.Range("K10").Value
            dados.Add "VENCIMENTO", ws.Range("L10").Value
            dados.Add "VOLUME", ws.Range("M10").Value
            dados.Add "DATA", ws.Range("N10").Value
            dados.Add "OPERA��O", ws.Range("O10").Value
            dados.Add "CAP", ws.Range("P10").Value

        Case "Rubi"
            dados.Add "ATIVO", ws.Range("J10").Value
            dados.Add "QUANTIDADE", ws.Range("K10").Value
            dados.Add "PRE�O REF", ws.Range("L10").Value
            dados.Add "VENCIMENTO", ws.Range("M10").Value
            dados.Add "STRIKE", ws.Range("N10").Value
            dados.Add "BARREIRA", ws.Range("O10").Value
            dados.Add "OPERA��O", ws.Range("P10").Value
        
        ' Caso nenhuma estrutura corresponda
        Case Else
            MsgBox "Estrutura n�o encontrada: " & nomeEstrutura
    End Select

    'Return da funcao
    CriarDict = dados
    Exit Function
End Function

Function montarEmail(emailAssessor As String, emailCliente As String, corpoEmail As String, assuntoEmail As String)
    ' Fun��o que monta o email de envio das opera��es
    
    'Args:
        '- emailAssessor: Email que ficar� em c�pia
        '- emailCliente: Email de quem receber� o email
        '- corpoEmail: Email

    Dim outApp As Outlook.Application
    Dim OutMail As Outlook.MailItem
    
    Set outApp = CreateObject("Outlook.Application")
    Set OutMail = outApp.CreateItem(olMailItem)

    With OutMail
        .Subject = assuntoEmail
        .HTMLBody = corpoEmail
        .CC = emailAssessor
        .To = emailCliente
        .Display
    End With

End Function

