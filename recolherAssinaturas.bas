Attribute VB_Name = "recolherAssinaturas"
Sub recolherAssinaturas()

    ' Marcar o in�cio do script
    Dim startTime As Double
    startTime = Timer
    
    'Definir email
    Dim email As String
    email = "datamanagementbrazil@whirlpool.com" 'Mudar este email
    
    'Definir senha
    Dim senha As String
    senha = "Rh2020" 'Mudar esta senha
    

    'Abrir o navegador e configurar a inst�ncia
    Dim bot As New ChromeDriver, by As New by 'Configura��es do Selenium
    bot.SetProfile Environ("localappdata") & "\Google\Chrome\User Data - Selenium" 'Diret�rio do profile 7 ( Selenium )
    bot.AddArgument "profile-directory=Profile 1" 'Acessar Profile 1
    bot.AddArgument "--disable-popup-blocking" 'Desabilita popups
    bot.AddArgument "--disable-notifications" 'Desabilita notifica��es
    bot.AddArgument "--disable-infobars" 'Desabilita infobars
    bot.Get "https://apps.docusign.com/send/home" 'Acessar p�gina principal do Docusign
    Application.Wait Now + TimeValue("00:00:05") ' Aguarde a p�gina carregar
    
    
    'Verifica se necessita de email para login
    If bot.IsElementPresent(by.Css("input[placeholder='Enter email']")) Then
        '''>>>ALTERAR EMAIL<<<'''
        bot.FindElementByName("email").SendKeys (email) & bot.keys.Enter
        'Aguardar a p�gina carregar
        Application.Wait Now + TimeValue("00:00:02")
    Else
        Debug.Print "O email n�o foi solicitado"
    End If
        'Verifica se necessita de senha para o login
    If bot.IsElementPresent(by.Css("input[placeholder='Enter password']")) Then
        '''>>>ALTERAR SENHA<<<'''
        bot.FindElementByName("password").SendKeys (senha) & bot.keys.Enter
        'Aguardar a p�gina carregar
        Application.Wait Now + TimeValue("00:00:02")
    Else
        Debug.Print "A senha n�o foi solicitada"
    End If
    ' Verifica se necessita de c�digo para verifica��o do email
    If bot.IsElementPresent(by.Css("input[placeholder='Enter code']")) Then
        ' Exibe a caixa de di�logo de entrada com o texto personalizado
        Dim codigoVerificacao As String
        codigoVerificacao = InputBox("Insira o c�digo de verifica��o enviado via e-mail e clique em continuar", "Verifica��o de C�digo")
    
        ' Verifica se o usu�rio inseriu um c�digo
        If codigoVerificacao <> "" Then
            ' C�digo para continuar com a macro
            Debug.Print "C�digo inserido: " & codigoVerificacao
            bot.FindElementByCss("input[placeholder='Enter code']").SendKeys codigoVerificacao ' Inserir o c�digo de verifica��o
            bot.SendKeys bot.keys.Enter 'Prosseguir para pr�xima p�gina
            ' Aguardar a p�gina carregar
            Application.Wait Now + TimeValue("00:00:02")
        Else
            ' C�digo para encerrar a macro se nenhum c�digo foi inserido
            Debug.Print "Nenhum c�digo inserido. Encerrando a macro."
        End If
    Else
        Debug.Print "O c�digo n�o foi solicitado"
    End If
    
    ' LOOP 0 - para acessar a aba "Gerenciar"
    Dim tentativas As Integer
    tentativas = 99  ' N�mero de tentativas

    Do While tentativas > 0
        On Error Resume Next
        bot.FindElementByXPath("(//span[normalize-space()='Gerenciar'])[1]").Click
        bot.FindElementByXPath("(//span[@class='menu_text css-gg4vpm'][normalize-space()='Conclu�do'])[1]").Click
        Application.Wait Now + TimeValue("00:00:02")  ' Ajuste conforme necess�rio
        On Error GoTo 0

        ' Verifique se a a��o foi bem-sucedida
        If Not Err.Number <> 0 Then
            Exit Do  ' Saia do loop se o clique for bem-sucedido
        End If

        ' Aguarde por um curto per�odo antes de tentar novamente
        Application.Wait Now + TimeValue("00:00:02")  ' Ajuste conforme necess�rio
        tentativas = tentativas - 1
    Loop


    ' Defina a planilha desejada
    Dim planilha As Worksheet
    Set planilha = ThisWorkbook.Sheets("Formul�rio")

    ' Defina o range de c�lulas A2:H
    Dim rng As Range
    Set rng = planilha.Range("A2:H" & planilha.Cells(planilha.Rows.Count, "A").End(xlUp).Row)
    
    
    ' Cole��o para armazenar linhas para o hist�rico
    Dim linhasParaHistorico As Collection
    Set linhasParaHistorico = New Collection

    ' Iterar sobre cada linha do range
    Dim cell As Range
    For Each cell In rng.Rows
        ' Verificar se a Assinatura j� foi recuperada na coluna I
        If cell.Cells(1, 9).Value <> "Assinatura recuperada" Then
            ' Obtenha os valores das c�lulas A, B, C, D, E, F e G para a linha atual
            Dim valorRE As String 'A
            Dim valorID As String 'B
            Dim valorNomeDest As String 'C
            Dim valorEmailDest As String 'D
            Dim valorNomeCC As String 'E
            Dim valorEmailCC As String 'F
            Dim valorAbsPath As Variant 'G
            Dim valorDataEnvio As String 'H
    
            valorRE = cell.Cells(1, 1).Value 'A
            valorID = cell.Cells(1, 2).Value 'B
            valorNomeDest = cell.Cells(1, 3).Value 'C
            valorEmailDest = cell.Cells(1, 4).Value 'D
            valorNomeCC = cell.Cells(1, 5).Value 'E
            valorEmailCC = cell.Cells(1, 6).Value 'F
            valorAbsPath = cell.Cells(1, 7).Value 'G
            valorDataEnvio = CDate(Left(cell.Cells(1, 8), 10)) 'H - Pegando os primeiros 10 digitos e transformando em Data
            
    
            'Procura pelo valorRE ( A ) dentro da caixa de pesquisa
            bot.FindElementByXPath("(//input[@placeholder='Pesquisar nas Visualiza��es r�pidas'])[1]").ClickDouble
            bot.FindElementByXPath("(//input[@placeholder='Pesquisar nas Visualiza��es r�pidas'])[1]").SendKeys valorRE
            bot.SendKeys bot.keys.Enter
            Application.Wait Now + TimeValue("00:00:02") ' Aguarde a p�gina carregar
            
        
            ' Procurar o valorRE em todos os elementos //div[@class='u-ellipsis']
            Dim elementosUellipsis As Object
            Set elementosUellipsis = bot.FindElementsByXPath("//div[@class='u-ellipsis']")
            
            ' Verificar se existem elementos u-ellipsis
            If elementosUellipsis.Count > 0 Then
                Dim i As Integer
                ' LOOP 2 - sobre os elementos da p�gina para encontrar o aviso correto
                For i = 1 To elementosUellipsis.Count
                    ' Atribuir o elemento a tituloDoEmail
                    Set tituloDoEmail = elementosUellipsis(i)
                
                    ' Verificar se o elemento span[@class='table_date'] est� presente para o mesmo �ndice
                    If bot.IsElementPresent(by.XPath("(//span[@class='table_date'])[" & i & "]")) Then
                        ' Atribuir o elemento a ultimaAltera��o
                        Set ultimaAltera��o = bot.FindElementByXPath("(//span[@class='table_date'])[" & i & "]")
                        teste = ultimaAltera��o.Text
                
                        ' Verificar se o t�tulo do e-mail cont�m valorRE
                        If InStr(1, tituloDoEmail.Text, valorRE, vbTextCompare) > 0 Then
                            ' Verificar se a �ltima altera��o � posterior � data limite
                            If CDate(ultimaAltera��o.Text) >= valorDataEnvio Then
                                ' Restante do c�digo permanece inalterado
                                Debug.Print "Para o caso " & valorRE & " o t�tulo do e-mail cont�m valorRE e a �ltima altera��o � posterior a " & Format(valorDataEnvio, "DD/MM/YYYY")
                                'Clicar em Download
                                bot.FindElementByXPath("(//button[@class='olv-button olv-ignore-transform css-lzi6zi'])[" & i & "]").Click
                                'Clicar em Todos
                                bot.FindElementByXPath("(//span[@class='css-1cbszci'][normalize-space()='Todos'])[1]").Click
                                'Clicar em Document
                                bot.FindElementByXPath("(//span[contains(text(),'Documento')])[1]").Click
                                'Clicar em Download
                                bot.FindElementByXPath("(//button[@class='olv-button olv-ignore-transform css-8evvwi'])[1]").Click
                
                                ' Esperar at� que o arquivo seja baixado (ajuste o caminho e o padr�o do nome do arquivo conforme necess�rio)
                                Dim filePath As String
                                Dim fileNamePattern As String
                                Dim timeout As Integer
                
                                ' Configura��es para esperar o download
                                filePath = "C:\Users\CAMBIA3\Documents\Selenium - Downloads\"
                                timeout = 60 ' Tempo m�ximo de espera em segundos
                                fileNamePattern = valorRE & "*"
                
                                If WaitForFile(filePath, fileNamePattern, timeout) Then
                                    ' O arquivo foi baixado com sucesso
                                    Debug.Print "Para o caso " & valorRE & " o download foi conclu�do!"
                                    ' Aqui voc� pode adicionar o restante do seu c�digo para lidar com o arquivo baixado, se necess�rio.
                                Else
                                    ' Tempo de espera excedido ou o arquivo n�o foi encontrado
                                    Debug.Print "Tempo de espera excedido ou arquivo n�o encontrado."
                                End If
                
                                'Informar se o aviso foi assinado e recuperado
                                cell.Cells(1, 9).Value = "Assinatura recuperada"
                                ' Adicionar a data e hora na coluna I
                                cell.Cells(1, 10).Value = Format(Now, "dd/mm/yyyy hh:mm:ss")
                
                                ' Verificar se a Assinatura foi recuperada
                                If cell.Cells(1, 9).Value = "Assinatura recuperada" Then
                                    ' Adicionar a linha � cole��o para mov�-la para o hist�rico posteriormente
                                    linhasParaHistorico.Add cell.EntireRow.Value
                                End If
                
                                ' Sair do loop interno, pois o elemento correto foi processado
                                Exit For
                            Else
                                Debug.Print "Para o caso " & valorRE & " o t�tulo do e-mail cont�m valorRE, mas a �ltima altera��o � anterior ou igual a " & Format(valorDataEnvio, "DD/MM/YYYY")
                                Application.Wait Now + TimeValue("00:00:02") ' Delay para preenchimento da planilha
                                cell.Cells(1, 9).Value = "Assinatura n�o recuperada"
                            End If
                        Else
                            Debug.Print "Para o caso " & valorRE & " o t�tulo do e-mail n�o cont�m valorRE"
                            Application.Wait Now + TimeValue("00:00:02") ' Delay para preenchimento da planilha
                            cell.Cells(1, 9).Value = "Assinatura n�o recuperada"
                        End If
                    Else
                        Debug.Print "Para o caso " & valorRE & " o elemento ultimaAltera��o n�o est� presente"
                        Application.Wait Now + TimeValue("00:00:02") ' Delay para preenchimento da planilha
                        cell.Cells(1, 9).Value = "Assinatura n�o recuperada"
                    End If
                Next i
            Else
                ' Caso o tituloDoEmail n�o for encontrado
                Debug.Print "Para o caso " & valorRE & " o elemento tituloDoEmail n�o est� presente"
                cell.Cells(1, 9).Value = "Assinatura n�o recuperada"
            End If
        End If
        Next cell
                
        ' Imprimir informa��es para depura��o
        Debug.Print "N�mero de linhas para o hist�rico: " & linhasParaHistorico.Count
        
        ' Mover linhas para o hist�rico
        Dim historicoSheet As Worksheet
        Set historicoSheet = ThisWorkbook.Sheets("Hist�rico")
        
        ' Inserir linhas na aba "Hist�rico"
        Dim linhaHistorico As Variant
        For Each linhaHistorico In linhasParaHistorico
            historicoSheet.Rows(historicoSheet.Cells(historicoSheet.Rows.Count, "A").End(xlUp).Row + 1).Value = linhaHistorico
        Next linhaHistorico
        Debug.Print "Linhas movidas para o hist�rico com sucesso."
       
       ' Excluir linhas com "Assinatura recuperada" na coluna I
        For i = rng.Rows.Count To 1 Step -1
            If planilha.Cells(i, 9).Value = "Assinatura recuperada" Then
                planilha.Rows(i).Delete
            End If
        Next i
        Debug.Print "Linhas com 'Assinatura recuperada' exclu�das com sucesso."
        
        ' Calcular o tempo decorrido em minutos
        Dim endTime As Double
        endTime = Timer
        Dim elapsedTime As Double
        elapsedTime = (endTime - startTime) / 60
        
        
        ' Exibir caixa de di�logo com o tempo decorrido e a mensagem de conclus�o
        Dim message As String
        message = "O script foi conclu�do em " & Format(elapsedTime, "0.00") & " minutos."
        MsgBox message, vbInformation, "Conclus�o da Tarefa"

End Sub
Function WaitForFile(filePath As String, fileNamePattern As String, timeout As Integer) As Boolean
    Dim startTime As Double
    Dim fso As Object
    Dim folder As Object
    Dim files As Object
    Dim file As Object
    Dim regex As Object
    Dim latestFile As Object
    Dim latestDate As Date

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = fileNamePattern

    Set folder = fso.GetFolder(filePath)
    startTime = Timer

    Do While Timer < startTime + timeout
        Set files = folder.files
        Set latestFile = Nothing
        latestDate = DateValue("01/01/1900")

        For Each file In files
            ' [INATIVO]  Log para verificar a an�lise dos arquivos
            'Debug.Print file
            If regex.Test(file.Name) Then
                ' Verificar se o arquivo � mais recente
                If file.DateLastModified > latestDate Then
                    Set latestFile = file
                    latestDate = file.DateLastModified
                End If
            End If
        Next file

        If Not latestFile Is Nothing Then
            ' O arquivo mais recente foi encontrado
            Debug.Print "Arquivo encontrado: " & latestFile.Name
            WaitForFile = True
            Exit Function
        End If

        ' Aguardar por um curto per�odo antes de verificar novamente
        Application.Wait Now + TimeValue("00:00:01")
    Loop

    ' Tempo de espera excedido
    WaitForFile = False
End Function
