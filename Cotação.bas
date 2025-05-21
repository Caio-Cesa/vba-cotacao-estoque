Attribute VB_Name = "Cotação"
Sub Cotacao_Fernando()
'Planilha no qual o programa está
Dim base As Worksheet
'Planilha de relatorio da linera emitido
Dim dado As Worksheet
'Planilha no qual sera alimentada as informações
Dim alvo As Worksheet
Dim cota As Worksheet
'Salva o nome da folha da plan em questão
Dim base2 As String
Dim dado2 As String
Dim alvo2 As String
Dim cota2 As String
Dim est As String
Dim ped As String
Dim prod As String
Dim cod As String
Dim emb As String
Dim filename As String
Dim currentPath As String
base2 = ActiveSheet.Name
Set base = Sheets(base2)
plan = base.Range("B2").Value
filename = "00dado"
currentPath = "C:\Users\Usuario\Desktop\"
''Vai abrir o documento de relatorio da linear com nome "dado"
'Workbooks.Open filename:=currentPath & filename & ".xls"
'dado2 = ActiveSheet.Name
'Set dado = Sheets(dado2)
''Exclusao de linhas unidas atrapalahndo a seleção da area da proxima fase
'dado.Range("A1:A11").EntireRow.Delete
'dado.Range("A65536").End(xlUp).EntireRow.Delete
''Metodo que transforma valores de texto em numero
'[A:B].Select
'With Selection
'    .NumberFormat = "General"
'    .Value = .Value
'End With
'posso abrir ja uma existente ou criar uma do zero mas colocando toda a formatação dps
Workbooks.Open filename:=currentPath & "COTAÇÃO.xls"
cota2 = ActiveSheet.Name
Set cota = Sheets(cota2)
linha_fimcota = cota.Range("C50000").End(xlUp).Row
cota.Range("A" & linha_fimcota + 1).Select
Workbooks.Open filename:=currentPath & plan
alvo2 = ActiveSheet.Name
Set alvo = Sheets(alvo2)
'precisa pegar col do pedido
col_fim = alvo.Range("XAA1").End(xlToLeft).Column + 1
linha_fim = alvo.Range("B50000").End(xlUp).Row
'cod = alvo.Cells.Find("COD").Offset(1, 0).Address(False, False, xlA1)
If alvo.Range(Cells(2, col_fim), Cells(2, col_fim)).Value = "PED." Then
    alvo.Range(Cells(3, col_fim), Cells(3, col_fim)).Select
    Do While ActiveCell.Row <= linha_fim
        alvo.Activate
        If ActiveCell.Value = "" Then
            ActiveCell.Offset(1).Select
        Else
            'IsEmpty (prod)
            alvo.Activate
            prod = ActiveCell.Offset(0, 2 - col_fim).Value
            cod = ActiveCell.Offset(0, 3 - col_fim).Value
            'emb = ActiveCell.Offset(0, 4 - col_fim).Value 'talvez seja quantidade nao tenho certeza ainda
            est = ActiveCell.Offset(0, -1).Value
            ped = ActiveCell.Value
            ActiveCell.Offset(1).Select
            cota.Activate
            ActiveCell.Value = est
            ActiveCell.Offset(0, 1).Value = ped
            ActiveCell.Offset(0, 2).Value = prod
            ActiveCell.Offset(0, 3).Value = cod
            'ActiveCell.Offset(0, 7).Value = emb
            ActiveCell.Offset(1).Select
        End If
    Loop
End If
Windows(plan).Close
'Formatação condicional Soma total
Range("A2:V10000").Select
Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=SE($B2=""Vencidos / a vencer"";1;0)"
Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
With Selection.FormatConditions(1).Font
    .Bold = True
    .Italic = False
    .Underline = xlUnderlineStyleSingle
    .TintAndShade = 0
End With
Selection.FormatConditions(1).StopIfTrue = False
End Sub
