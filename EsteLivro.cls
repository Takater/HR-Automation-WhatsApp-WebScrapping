VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EsteLivro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub Workbook_Activate()

    'Checar data do �ltimo envio
    Do While Not Sheets("HOJE").Cells(2, 6) = Date
    
        'Mes atual
        Dim mes
        mes = StrConv(MonthName(Month(Date)), vbProperCase)
        
        'Contar quantos come�am HOJE
        Dim counter, row
        counter = 0
        row = 2
        Do
            If Sheets(mes).Cells(row, 11) = Date Then counter = counter + 1
            row = row + 1
            
        Loop Until IsEmpty(Sheets(mes).Cells(row, 11))
        
        'Se houverem pessoas iniciando
        If counter > 0 Then
        
            'Caixa de pergunta
            response = MsgBox((CStr(counter) + " pessoa(s) come�a(m) hoje. Enviar mensagens?"), vbYesNo, "Boas-vindas")
            
            'Se confirmar envio de mensagens
            If response = vbYes Then
                
                'Montar Lista
                MontarLista
                
                'Enviar Arquivos
                EnviarArq
                
                Exit Sub
                
            'Se n�o confirmar envio de mensagens
            Else
                
                'Sair sem atualizar data de envio
                Exit Sub
                
            End If
            
        'Se n�o houverem pessoas iniciando
        Else
        
            'Caixa de Mensagem
            response = MsgBox("N�o h� pessoas iniciando hoje", vbOKOnly, "Boas-Vindas")
            If response = vbOK Then
                
                'Data de envio atualizada
                Sheets("HOJE").Cells(2, 6) = Date
                
                Exit Sub
            End If
        End If
    Loop
End Sub
