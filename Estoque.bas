Attribute VB_Name = "Estoque"
Sub Estoque_Fernando()
'Planilha no qual o programa está
Dim base As Worksheet
'Planilha de relatorio da linera emitido
Dim dado As Worksheet
'Planilha no qual sera alimentada as informações
Dim alvo As Worksheet
'Salva o nome da folha da plan em questão
Dim base2 As String
Dim dado2 As String
Dim alvo2 As String
Dim filename As String
Dim currentPath As String
base2 = ActiveSheet.Name
Set base = Sheets(base2)
plan = base.Range("B2").Value
filename = "00dado"
currentPath = "C:\Users\Usuario\Desktop\"
'Vai abrir o documento de relatorio de estoque com nome "00dado" FICAR A TENTO SE ARQUIVO É XLS OU XLSX
Workbooks.Open filename:=currentPath & filename & ".xls"
dado2 = ActiveSheet.Name
Set dado = Sheets(dado2)
dado.Range("A1:A11").EntireRow.Delete
dado.Range("A65536").End(xlUp).EntireRow.Delete
'Metodo que transforma valores de texto em numero
[A:B].Select
With Selection
    .NumberFormat = "General"
    .Value = .Value
End With
Workbooks.Open filename:=currentPath & plan
alvo2 = ActiveSheet.Name
Set alvo = Sheets(alvo2)
col_fim = alvo.Range("XAA1").End(xlToLeft).Column + 2
linha_fim = alvo.Range("B50000").End(xlUp).Row
cod = alvo.Cells.Find("COD").Offset(1, 0).Address(False, False, xlA1)
alvo.Range(Cells(1, col_fim), Cells(1, col_fim)).FormulaLocal = "=HOJE()"
'usando procv para puxar dados de estoque que fica na coluna 11 da plan 00dados sheet1
alvo.Range(Cells(3, col_fim), Cells(linha_fim / 2, col_fim)).FormulaLocal = "=SEERRO(PROCV(" & cod & ";[" & filename & ".xls]" & dado2 & "!$A:$AA;11;FALSO);"""")"
alvo.Range(Cells(1, col_fim), Cells(linha_fim / 2, col_fim)).Value = alvo.Range(Cells(1, col_fim), Cells(linha_fim / 2, col_fim)).Value
'seleciona a coluna no qual os dados foram colocados
alvo.Range(Cells(linha_fim, col_fim), Cells(linha_fim, col_fim)).Offset(1).Select
Application.DisplayAlerts = False
Windows(filename & ".xls").Close
Application.DisplayAlerts = True
'Daqui pra baixo é formatação da celulas editadas para o padrão
alvo.Range(Cells(1, col_fim), Cells(linha_fim, col_fim + 1)).ClearFormats
alvo.Range(Cells(3, col_fim), Cells(linha_fim, col_fim + 1)).NumberFormat = "00"
alvo.Range(Cells(1, col_fim), Cells(1, col_fim + 1)).NumberFormat = "dd/mm/yy"
alvo.Range(Cells(1, col_fim), Cells(1, col_fim + 1)).Merge
alvo.Range(Cells(1, col_fim), Cells(linha_fim, col_fim + 1)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlVAlignCenter
        .Orientation = 0
        .Font.Bold = True
        .Font.Size = 12
    End With
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
alvo.Range(Cells(2, col_fim), Cells(2, col_fim + 1)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
alvo.Range(Cells(2, col_fim), Cells(2, col_fim + 1)).Font.Size = 10
alvo.Range(Cells(2, col_fim), Cells(2, col_fim)).FormulaLocal = "EST."
alvo.Range(Cells(2, col_fim + 1), Cells(2, col_fim + 1)).FormulaLocal = "PED."
End Sub
