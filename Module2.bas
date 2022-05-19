Attribute VB_Name = "Module2"
Sub EnviarArq()
    
    'Variaveis
    Dim pagina As New Selenium.ChromeDriver
    Dim keys As New Selenium.keys
    Dim file, URL, user As String
    
    'Coookies do Chrome
    user = Left(Application.UserLibraryPath, 20)
    pagina.SetProfile user & "\AppData\Local\Google\Chrome\User Data", True
    pagina.AddArgument (user & "\AppData\Local\Google\Chrome\User Data\Default")
    pagina.AddArgument ("--no-sandbox")
    
    'Caso haja janela do Chrome aberta
    On Error GoTo TE
    
    'Arquivo de vídeo
    file = Sheets("HOJE").Cells(2, 5)
    
    'Dentro do Navegador
    With pagina
        .Timeouts.PageLoad = 100000
        .Start
        .Get ("http://web.whatsapp.com")
        Application.Wait (Now + TimeValue("00:00:20"))
    
        'Loop pela Lista
        Dim row As Integer
        row = 2
        Do While Not IsEmpty(Sheets("HOJE").Cells(row, 1))
            
            'Montar Variaveis
            Dim nome, phone, msg, part1, part2 As String
            
            With Sheets("HOJE")
                
                'Nome
                nome = StrConv(Split(.Cells(row, 1), " ")(0), vbProperCase)
                
                'Telefone
                phone = "55"
                For counter = 1 To Len(.Cells(row, 3))
                    If IsNumeric(Mid(.Cells(row, 3), counter, 1)) Then phone = phone & Mid(.Cells(row, 3), counter, 1)
                Next
                
                'Mensagem
                If StrComp(Left(.Cells(row, 2), 1), "M") = 0 Then part1 = "Seja+bem-vindo%2C+" Else part1 = "Seja+bem-vinda%2C+"
                part2 = "%21+%0D%0AFicamos+muito+felizes+com+sua+vinda+para+o+nosso+time%21%0D%0ADesejamos+muito+sucesso+nesta+nova+trajet%C3%B3ria.%0D%0AConte+sempre+conosco.+%F0%9F%98%8A"
                msg = part1 & nome & part2
                
            End With
            
            'Montar URL
            URL = "https://api.whatsapp.com/send/?phone=" & phone & "&text=" & msg
    
            'Abrir conversa
            .Get (URL)
            .FindElementByXPath("/html/body/div[1]/div[1]/div[2]/div/section/div/div/div/div[2]/div[1]/a").Click
            Application.Wait (Now + TimeValue("00:00:10"))
            
            'Selecionar Arquivo e enviar mensagem
            .FindElementByXPath("/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div").Click
            Application.Wait (Now + TimeValue("00:00:02"))
            .FindElementByCss("input[type='file']").SendKeys file
            Application.Wait (Now + TimeValue("00:00:10"))
            .FindElementByXPath("/html/body/div[1]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div").Click
            Application.Wait (Now + TimeValue("00:00:10"))
            
            'Anotar Erro
            Sheets("HOJE").Cells(row, 4).Value = "Falha"
            On Error GoTo ErrorChrome
            
            'Confirmar Envio
            Sheets("HOJE").Cells(row, 4).Value = "Sucesso"
            
ErrorChrome:
            'Proximo Contato
            row = row + 1
        Loop
    End With
    Sheets("HOJE").Cells(2, 6) = Date
    
TE:
    TerminateChrome
    EnviarArq
    Exit Sub
End Sub

Public Sub TerminateChrome()
    Dim objWMIcimv2         As Object
    Dim objProcess          As Object
    Dim objProcesses        As Object
    Dim lngError            As Long

    Const strTerminateThis  As String = "chrome.exe"

    Set objWMIcimv2 = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set objProcesses = objWMIcimv2.ExecQuery("select * from win32_process where name='" & strTerminateThis & "'")

    For Each objProcess In objProcesses
        lngError = objProcess.Terminate
        If lngError = 0 Then Exit For
    Next

    Set objWMIcimv2 = Nothing
    Set objProcesses = Nothing
    Set objProcess = Nothing
End Sub
