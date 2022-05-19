Attribute VB_Name = "Module1"
Sub MontarLista()
    
    'Mes atual
    Dim mes
    mes = StrConv(MonthName(Month(Date)), vbProperCase)
    
    
    'Limpar linhas da tabela de destino
    Dim row
    row = 2
    
    Do
        'Limpar colunas de A at� C
        With Folha4
            Range(.Cells(row, 1), .Cells(row, 3)).ClearContents
        End With
        
        'Ir para pr�xima linha
        row = row + 1
        
    Loop Until IsEmpty(Sheets("HOJE").Cells(row, 1))
    
    
    'Checar colaboradores e montar lista
    Dim i, j
    i = 2
    j = 2
    
    Do
        'Data de �nicio igual a HOJE
        If Sheets(mes).Cells(i, 11) = Date Then
        
            'Nome
            Sheets("HOJE").Cells(j, 1) = Sheets(mes).Cells(i, 5)
            
            'Gen�ro
            Sheets("HOJE").Cells(j, 2) = Sheets(mes).Cells(i, 6)
            
            'N�mero de Telefone
            Sheets("HOJE").Cells(j, 3) = Sheets(mes).Cells(i, 34)
            
            'Pr�xima linha da tabela de destino
            j = j + 1
            
        End If
        
        'Pr�xima linha da tabela de origem
        i = i + 1
        
    Loop Until IsEmpty(Sheets(mes).Cells(i, 11))
    
End Sub
