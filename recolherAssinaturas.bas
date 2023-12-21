Attribute VB_Name = "recolherAssinaturas"
Sub recolherAssinaturas()

    ' Marcar o início do script
    Dim startTime As Double
    startTime = Timer
    
    'Definir email
    Dim email As String
    email = "datamanagementbrazil@whirlpool.com" 'Mudar este email
    
    'Definir senha
    Dim senha As String
    senha = "Rh2020" 'Mudar esta senha
    

    'Abrir o navegador e configurar a instância
    Dim bot As New ChromeDriver, by As New by 'Configurações do Selenium
    bot.SetProfile Environ("localappdata") & "\Google\Chrome\User Data - Selenium" 'Diretório do profile 7 ( Selenium )
    bot.AddArgument "profile-directory=Profile 1" 'Acessar Profile 1
    bot.AddArgument "--disable-popup-blocking" 'Desabilita popups
    bot.AddArgument "--disable-notifications" 'Desabilita notificações
    bot.AddArgument "--disable-infobars" 'Desabilita infobars
    bot.Get "https://apps.docusign.com/send/home" 'Acessar página principal do Docusign
    Application.Wait Now + TimeValue("00:00:05") ' Aguarde a página carregar
    
    
    'Verifica se necessita de email para login
    If bot.IsElementPresent(by.Css("input[placeholder='Enter email']")) Then
        '''>>>ALTERAR EMAIL<<<'''
        bot.FindElementByName("email").SendKeys (email) & bot.keys.Enter
        'Aguardar a página carregar
        Application.Wait Now + TimeValue("00:00:02")
    Else
        Debug.Print "O email não foi solicitado"
    End If
        'Verifica se necessita de senha para o login
    If bot.IsElementPresent(by.Css("input[placeholder='Enter password']")) Then
        '''>>>ALTERAR SENHA<<<'''
        bot.FindElementByName("password").SendKeys (senha) & bot.keys.Enter
        'Aguardar a página carregar
        Application.Wait Now + TimeValue("00:00:02")
    Else
        Debug.Print "A senha não foi solicitada"
    End If
    ' Verifica se necessita de código para verificação do email
    If bot.IsElementPresent(by.Css("input[placeholder='Enter code']")) Then
        ' Exibe a caixa de diálogo de entrada com o texto personalizado
        Dim codigoVerificacao As String
        codigoVerificacao = InputBox("Insira o código de verificação enviado via e-mail e clique em continuar", "Verificação de Código")
    
        ' Verifica se o usuário inseriu um código
        If codigoVerificacao <> "" Then
            ' Código para continuar com a macro
            Debug.Print "Código inserido: " & codigoVerificacao
            bot.FindElementByCss("input[placeholder='Enter code']").SendKeys codigoVerificacao ' Inserir o código de verificação
            bot.SendKeys bot.keys.Enter 'Prosseguir para próxima página
            ' Aguardar a página carregar
            Application.Wait Now + TimeValue("00:00:02")
        Else
            ' Código para encerrar a macro se nenhum código foi inserido
            Debug.Print "Nenhum código inserido. Encerrando a macro."
        End If
    Else
        Debug.Print "O código não foi solicitado"
    End If
    
    ' LOOP 0 - para acessar a aba "Gerenciar"
    Dim tentativas As Integer
    tentativas = 99  ' Número de tentativas

    Do While tentativas > 0
        On Error Resume Next
        bot.FindElementByXPath("(//span[normalize-space()='Gerenciar'])[1]").Click
        bot.FindElementByXPath("(//span[@class='menu_text css-gg4vpm'][normalize-space()='Concluído'])[1]").Click
        Application.Wait Now + TimeValue("00:00:02")  ' Ajuste conforme necessário
        On Error GoTo 0

        ' Verifique se a ação foi bem-sucedida
        If Not Err.Number <> 0 Then
            Exit Do  ' Saia do loop se o clique for bem-sucedido
        End If

        ' Aguarde por um curto período antes de tentar novamente
        Application.Wait Now + TimeValue("00:00:02")  ' Ajuste conforme necessário
        tentativas = tentativas - 1
    Loop


    ' Defina a planilha desejada
    Dim planilha As Worksheet
    Set planilha = ThisWorkbook.Sheets("Formulário")

    ' Defina o range de células A2:H
    Dim rng As Range
    Set rng = planilha.Range("A2:H" & planilha.Cells(planilha.Rows.Count, "A").End(xlUp).Row)
    
    
    ' Coleção para armazenar linhas para o histórico
    Dim linhasParaHistorico As Collection
    Set linhasParaHistorico = New Collection

    ' Iterar sobre cada linha do range
    Dim cell As Range
    For Each cell In rng.Rows
        ' Verificar se a Assinatura já foi recuperada na coluna I
        If cell.Cells(1, 9).Value <> "Assinatura recuperada" Then
            ' Obtenha os valores das células A, B, C, D, E, F e G para a linha atual
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
            bot.FindElementByXPath("(//input[@placeholder='Pesquisar nas Visualizações rápidas'])[1]").ClickDouble
            bot.FindElementByXPath("(//input[@placeholder='Pesquisar nas Visualizações rápidas'])[1]").SendKeys valorRE
            bot.SendKeys bot.keys.Enter
            Application.Wait Now + TimeValue("00:00:02") ' Aguarde a página carregar
            
        
            ' Procurar o valorRE em todos os elementos //div[@class='u-ellipsis']
            Dim elementosUellipsis As Object
            Set elementosUellipsis = bot.FindElementsByXPath("//div[@class='u-ellipsis']")
            
            ' Verificar se existem elementos u-ellipsis
            If elementosUellipsis.Count > 0 Then
                Dim i As Integer
                ' LOOP 2 - sobre os elementos da página para encontrar o aviso correto
                For i = 1 To elementosUellipsis.Count
                    ' Atribuir o elemento a tituloDoEmail
                    Set tituloDoEmail = elementosUellipsis(i)
                
                    ' Verificar se o elemento span[@class='table_date'] está presente para o mesmo índice
                    If bot.IsElementPresent(by.XPath("(//span[@class='table_date'])[" & i & "]")) Then
                        ' Atribuir o elemento a ultimaAlteração
                        Set ultimaAlteração = bot.FindElementByXPath("(//span[@class='table_date'])[" & i & "]")
                        teste = ultimaAlteração.Text
                
                        ' Verificar se o título do e-mail contém valorRE
                        If InStr(1, tituloDoEmail.Text, valorRE, vbTextCompare) > 0 Then
                            ' Verificar se a última alteração é posterior à data limite
                            If CDate(ultimaAlteração.Text) >= valorDataEnvio Then
                                ' Restante do código permanece inalterado
                                Debug.Print "Para o caso " & valorRE & " o título do e-mail contém valorRE e a última alteração é posterior a " & Format(valorDataEnvio, "DD/MM/YYYY")
                                'Clicar em Download
                                bot.FindElementByXPath("(//button[@class='olv-button olv-ignore-transform css-lzi6zi'])[" & i & "]").Click
                                'Clicar em Todos
                                bot.FindElementByXPath("(//span[@class='css-1cbszci'][normalize-space()='Todos'])[1]").Click
                                'Clicar em Document
                                bot.FindElementByXPath("(//span[contains(text(),'Documento')])[1]").Click
                                'Clicar em Download
                                bot.FindElementByXPath("(//button[@class='olv-button olv-ignore-transform css-8evvwi'])[1]").Click
                
                                ' Esperar até que o arquivo seja baixado (ajuste o caminho e o padrão do nome do arquivo conforme necessário)
                                Dim filePath As String
                                Dim fileNamePattern As String
                                Dim timeout As Integer
                
                                ' Configurações para esperar o download
                                filePath = "C:\Users\CAMBIA3\Documents\Selenium - Downloads\"
                                timeout = 60 ' Tempo máximo de espera em segundos
                                fileNamePattern = valorRE & "*"
                
                                If WaitForFile(filePath, fileNamePattern, timeout) Then
                                    ' O arquivo foi baixado com sucesso
                                    Debug.Print "Para o caso " & valorRE & " o download foi concluído!"
                                    ' Aqui você pode adicionar o restante do seu código para lidar com o arquivo baixado, se necessário.
                                Else
                                    ' Tempo de espera excedido ou o arquivo não foi encontrado
                                    Debug.Print "Tempo de espera excedido ou arquivo não encontrado."
                                End If
                
                                'Informar se o aviso foi assinado e recuperado
                                cell.Cells(1, 9).Value = "Assinatura recuperada"
                                ' Adicionar a data e hora na coluna I
                                cell.Cells(1, 10).Value = Format(Now, "dd/mm/yyyy hh:mm:ss")
                
                                ' Verificar se a Assinatura foi recuperada
                                If cell.Cells(1, 9).Value = "Assinatura recuperada" Then
                                    ' Adicionar a linha à coleção para movê-la para o histórico posteriormente
                                    linhasParaHistorico.Add cell.EntireRow.Value
                                End If
                
                                ' Sair do loop interno, pois o elemento correto foi processado
                                Exit For
                            Else
                                Debug.Print "Para o caso " & valorRE & " o título do e-mail contém valorRE, mas a última alteração é anterior ou igual a " & Format(valorDataEnvio, "DD/MM/YYYY")
                                Application.Wait Now + TimeValue("00:00:02") ' Delay para preenchimento da planilha
                                cell.Cells(1, 9).Value = "Assinatura não recuperada"
                            End If
                        Else
                            Debug.Print "Para o caso " & valorRE & " o título do e-mail não contém valorRE"
                            Application.Wait Now + TimeValue("00:00:02") ' Delay para preenchimento da planilha
                            cell.Cells(1, 9).Value = "Assinatura não recuperada"
                        End If
                    Else
                        Debug.Print "Para o caso " & valorRE & " o elemento ultimaAlteração não está presente"
                        Application.Wait Now + TimeValue("00:00:02") ' Delay para preenchimento da planilha
                        cell.Cells(1, 9).Value = "Assinatura não recuperada"
                    End If
                Next i
            Else
                ' Caso o tituloDoEmail não for encontrado
                Debug.Print "Para o caso " & valorRE & " o elemento tituloDoEmail não está presente"
                cell.Cells(1, 9).Value = "Assinatura não recuperada"
            End If
        End If
        Next cell
                
        ' Imprimir informações para depuração
        Debug.Print "Número de linhas para o histórico: " & linhasParaHistorico.Count
        
        ' Mover linhas para o histórico
        Dim historicoSheet As Worksheet
        Set historicoSheet = ThisWorkbook.Sheets("Histórico")
        
        ' Inserir linhas na aba "Histórico"
        Dim linhaHistorico As Variant
        For Each linhaHistorico In linhasParaHistorico
            historicoSheet.Rows(historicoSheet.Cells(historicoSheet.Rows.Count, "A").End(xlUp).Row + 1).Value = linhaHistorico
        Next linhaHistorico
        Debug.Print "Linhas movidas para o histórico com sucesso."
       
       ' Excluir linhas com "Assinatura recuperada" na coluna I
        For i = rng.Rows.Count To 1 Step -1
            If planilha.Cells(i, 9).Value = "Assinatura recuperada" Then
                planilha.Rows(i).Delete
            End If
        Next i
        Debug.Print "Linhas com 'Assinatura recuperada' excluídas com sucesso."
        
        ' Calcular o tempo decorrido em minutos
        Dim endTime As Double
        endTime = Timer
        Dim elapsedTime As Double
        elapsedTime = (endTime - startTime) / 60
        
        
        ' Exibir caixa de diálogo com o tempo decorrido e a mensagem de conclusão
        Dim message As String
        message = "O script foi concluído em " & Format(elapsedTime, "0.00") & " minutos."
        MsgBox message, vbInformation, "Conclusão da Tarefa"

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
            ' [INATIVO]  Log para verificar a análise dos arquivos
            'Debug.Print file
            If regex.Test(file.Name) Then
                ' Verificar se o arquivo é mais recente
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

        ' Aguardar por um curto período antes de verificar novamente
        Application.Wait Now + TimeValue("00:00:01")
    Loop

    ' Tempo de espera excedido
    WaitForFile = False
End Function
