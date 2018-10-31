# vba
# Macro que verifica se cell "A?" é igual a cell"B1", sendo verdadeiro copia o valor de "B?" para "C?", se o valor for vazio, ela buscara
# o ultimo valor valido correspondete 

Sub Macro_Prenche()

    Dim i As Integer
    Dim Qtde As Integer
    Dim av As Integer
    
    Qtde = Cells(Rows.Count, "A").End(xlUp).Row
     
    For i = 2 To Qtde
        
        If Plan1.Cells(i, 1).Value = Plan1.Cells(1, 2).Value Then
        av = Plan1.Cells(i, 2).Value
            
            If av = Empty Then
            lRow = Plan1.Cells(i, 2).End(xlUp).Row
            av = Plan1.Cells(lRow, 2).Value
            End If
            
        Plan1.Cells(i, 3).Select
        ActiveCell.FormulaR1C1 = av
        End If
        
      
    Next i
    
     
End Sub
