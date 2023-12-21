Attribute VB_Name = "enviarAvisos"
Sub enviarAvisoDeFerias()

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
    bot.Get "https://apps.docusign.com/send/documents?type=envelopes" 'Acessar p�gina principal do Docusign
    Application.Wait Now + TimeValue("00:00:02") ' Aguarde a p�gina carregar
    
    
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

    ' Defina a planilha desejada
    Dim planilha As Worksheet
    Set planilha = ThisWorkbook.Sheets("Formul�rio")

    ' Defina o range de c�lulas A2:H
    Dim rng As Range
    Set rng = planilha.Range("A2:H" & planilha.Cells(planilha.Rows.Count, "A").End(xlUp).Row)

    ' Iterar sobre cada linha do range
    Dim cell As Range
    For Each cell In rng.Rows
        ' Verificar se a coluna A n�o � nula e a coluna G est� vazia
        If Not IsEmpty(cell.Cells(1, 1).Value) And IsEmpty(cell.Cells(1, 8).Value) Then
        ' Obtenha os valores das c�lulas A, B, C, D, E, F e G para a linha atual
            Dim valorRE As String
            Dim valorID As String
            Dim valorNomeDest As String
            Dim valorEmailDest As String
            Dim valorNomeCC As String
            Dim valorEmailCC As String
            Dim valorAbsPath As Variant
    
            valorRE = cell.Cells(1, 1).Value 'A
            valorID = cell.Cells(1, 2).Value 'B
            valorNomeDest = cell.Cells(1, 3).Value 'C
            valorEmailDest = cell.Cells(1, 4).Value 'D
            valorNomeCC = cell.Cells(1, 5).Value 'E
            valorEmailCC = cell.Cells(1, 6).Value 'F
            valorAbsPath = cell.Cells(1, 7).Value 'G
        
        
        'Abrir template
        bot.FindElementByXPath("(//button[@class='olv-button olv-ignore-transform css-17ozirp'])[1]").Click
        ' Clicar em Enviar um envelope
        bot.FindElementByXPath("(//span[normalize-space()='Enviar um envelope'])[1]").Click
        
        'Fazer upload do arquivo
        'Definir o n�mero m�ximo de tentativas
        Const maxTentativas As Integer = 20
        
        'Realizar o upload
        Dim tentativas As Integer
        tentativas = 0
        
        Do While tentativas < maxTentativas
            On Error Resume Next
            'Aguardar o bot�o carregar para upload
            Application.Wait Now + TimeValue("00:00:01")
            'Enviar arquivo para upload
            bot.FindElementByClass("css-89bprp").SendKeys valorAbsPath
            On Error GoTo 0 ' Restaura o tratamento normal de erros
            
            'Verificar se ocorreu um erro durante o upload
            If Err.Number = 0 Then
                'Verificar se o elemento est� presente ap�s o upload
                If bot.IsElementPresent(by.XPath("(//div[@class='css-36j3xk'])[1]")) Then
                    Debug.Print "O elemento est� presente"
                    nomeDoArquivo = bot.FindElementByXPath("(//div[@class='css-36j3xk'])[1]").Text
                    
                    'Verificar se o upload foi bem-sucedido
                    If InStr(1, nomeDoArquivo, valorRE, vbTextCompare) > 0 Then
                        Debug.Print "O upload foi realizado com sucesso"
                        Exit Do ' Saia do loop se o upload for bem-sucedido
                    End If
                Else
                    Debug.Print "O elemento n�o est� presente"
                End If
            Else
                'C�digo para lidar com o erro
                MsgBox "Ocorreu um erro: " & Err.Description
            End If
            
            'Incrementar o n�mero de tentativas
            tentativas = tentativas + 1
            Debug.Print "N�mero de tentativas de upload:"; tentativas
            
            'Aguardar um curto per�odo antes de tentar novamente
            Application.Wait Now + TimeValue("00:00:01")
        Loop
        
        'Verificar se o n�mero m�ximo de tentativas foi atingido
        If tentativas = maxTentativas Then
            MsgBox "N�mero m�ximo de tentativas atingido. O upload n�o foi bem-sucedido."
        End If
                
                
                ' LOOP INICIO #Janela Cancelar - para aguardar a janela de modelos abrir e clicar em cancelar
                    tentativas = 20 ' N�mero de tentativas
                    
                    Do While tentativas > 0
                        On Error Resume Next
                        ' Procurar a caixaDeTemplate
                        Dim caixaDeTemplate As Object
                        Set caixaDeTemplate = bot.FindElementsByXPath("(//span[normalize-space()='Selecione os modelos correspondentes'])[1]")
                        
                        ' Verificar se existem elementos caixaDeTemplate
                        If caixaDeTemplate.Count > 0 Then
                            bot.FindElementByXPath("(//button[@class='olv-button olv-ignore-transform css-mtra3x'])[1]").Click
                            Application.Wait Now + TimeValue("00:00:01")  ' Ajuste conforme necess�rio
                            On Error GoTo 0
                    
                            ' Verifique se a a��o foi bem-sucedida
                            If Err.Number = 0 Then
                                Debug.Print "O pop-up de templates foi fechado"
                                Exit Do  ' Saia do loop se o clique for bem-sucedido
                            End If
                        Else
                            ' Aguarde por um curto per�odo antes de tentar novamente
                            Application.Wait Now + TimeValue("00:00:01")  ' Ajuste conforme necess�rio
                            tentativas = tentativas - 1
                            Debug.Print "N�mero de vezes que o script tentou cancelar a janela de modelos de templates:"; tentativas
                        End If
                    '  LOOP  FIM #Janela Cancelar - Realiza o upload
                    Loop
        
        'Preencher  nome do destinat�rio
        bot.FindElementByClass("css-1ez4hss").SendKeys valorNomeDest
        'Preencher email do destinat�rio
        bot.FindElementByClass("css-gbdw5j").SendKeys valorEmailDest
        bot.SendKeys bot.keys.Enter
        
        'Aguardar a p�gina carregar
        Application.Wait Now + TimeValue("00:00:01")  ' Ajuste conforme necess�rio
        
        'Clicar em Adicionar Destinat�rio
        bot.FindElementByXPath("(//button[@class='olv-button olv-ignore-transform css-1cn55yz'])[1]").Click
        'Preencher destinat�rio em c�pia
        bot.FindElementByXPath("(//input[contains(@role,'combobox')])[3]").SendKeys valorNomeCC
        'Aguardar a p�gina carregar
        Application.Wait Now + TimeValue("00:00:01")  ' Ajuste conforme necess�rio
        'Preencher email em c�pia
        bot.FindElementByXPath("(//input[contains(@role,'combobox')])[4]").SendKeys valorEmailCC
        'Aguardar a p�gina carregar
        Application.Wait Now + TimeValue("00:00:01")  ' Ajuste conforme necess�rio
        'Trocar op��o para apenas em c�pia
        bot.FindElementByXPath("(//button[@type='button'])[13]").Click
        bot.FindElementByCss("button[data-qa='carbonCopies']").Click
        
        'Preencher ID de Venda
        bot.FindElementByXPath("(//input[@data-qa='label-input-ID de venda'])[1]").SendKeys valorRE
        
        'Clicar no Tipo de Documento
        Dim dropDown As SelectElement
        'Abrir dropdown para selecionar a op��o correta
        Set dropDown = bot.FindElementByXPath("(//select[@class='css-12ihcxq'])[1]").AsSelect
        'Selecionar Aditivo Contratual
        dropDown.SelectByText ("Aditivo Contratual")
        
        
        'Clicar no campo de texto do email
        bot.FindElementByXPath("(//textarea[@placeholder='Inserir mensagem'])[1]").Click
        'Preenchimento do texto do campo de texto do email
        bot.SendKeys ("Prezado(a)")
        bot.SendKeys bot.keys.Enter 'Par�grafo
        bot.SendKeys bot.keys.Enter 'Par�grafo
        bot.SendKeys bot.keys.Enter 'Par�grafo
        bot.SendKeys ("O per�odo de f�rias foi programado conforme solicitado.")
        bot.SendKeys bot.keys.Enter 'Par�grafo
        bot.SendKeys ("Segue o aviso para assinatura.")
        bot.SendKeys bot.keys.Enter 'Par�grafo
        bot.SendKeys bot.keys.Enter 'Par�grafo
        bot.SendKeys ("Obs.: N�o � necess�ria a entrega do aviso f�sico.")
        bot.SendKeys bot.keys.Enter 'Par�grafo
        bot.SendKeys bot.keys.Enter 'Par�grafo
        bot.SendKeys bot.keys.Enter 'Par�grafo
        bot.SendKeys ("Atenciosamente,")
        bot.SendKeys bot.keys.Enter 'Par�grafo
        bot.SendKeys ("RH Whirlpool.")
               
               
        'Prosseguir para template
        bot.FindElementByXPath("(//button[@class='olv-button olv-ignore-transform css-1n4kelk'])[1]").Click
        
        'Diminuir o zoom da p�gina para colocar os campos de assinatura
        bot.FindElementByXPath("(//button[@class='css-zspncp'])[1]").Click
        bot.FindElementByXPath("(//span[normalize-space()='50%'])[1]").Click
        'Colocar o primeiro campo de assinatura
        bot.FindElementByXPath("(//span[@title='Assinatura'])[1]").ClickAndHold
        bot.Actions.MoveByOffset(270, 34).Release.Perform
        'Colocar o segundo campo de assinatura
        bot.FindElementByCss("span[title='Assinatura']").ClickAndHold
        bot.Actions.MoveByOffset(270, 240).Release.Perform
        
        'Prosseguir para envio
        bot.FindElementByXPath("(//button[@class='olv-button olv-ignore-transform css-1n4kelk'])[1]").Click
        'Aguardar o envio
        Application.Wait Now + TimeValue("00:00:02")  ' Ajuste conforme necess�rio
        
        ' Adicionar a data e hora na coluna D
        cell.Cells(1, 8).Value = Format(Now, "dd/mm/yyyy hh:mm:ss")
        Application.Wait Now + TimeValue("00:00:02")  ' Ajuste conforme necess�rio
        End If
    Next cell

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
