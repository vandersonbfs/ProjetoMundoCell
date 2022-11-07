Attribute VB_Name = "Módulo4"
Sub InserirFluxo()
bloqueado = True
    Dim tabela As ListObject
    Dim n As Integer
    Dim id As Integer
    Dim os As Long
    Dim hj As Date
    Dim Qntd As Integer
    Dim vlr As Long
    Dim Descr As String
    Dim tipo As String


    If Planilha6.Range("L2").Value = "" Then
        id = Range("OS").Value 'Recebe numero novo de OS
        tipo = "ENTRADA"
        Descr = "VENDA"
    ElseIf id = slv Then
        os = Range("Slv").Value 'Se não, Recebe numero do serviço executado
        tipo = "ENTRADA"
        Descr = "SERVIÇO"
    End If

    Set tabela = Planilha5.ListObjects(1)
    n = tabela.Range.Rows.Count
    id = Range("IDcaixa").Value
   
    hj = Range("HJ").Value
    
    tabela.Range(n, 1).Value = id
    tabela.Range(n, 2).Value = hj
    tabela.Range(n, 3).Value = os
    tabela.Range(n, 4).Value = tipo
    tabela.Range(n, 5).Value = Descr
    tabela.Range(n, 6).Value = UserForm4.lblSubTotal
    
    Range("IDcaixa").Value = id + 1

    If Descr = "VENDA" Then
            Range("OS").Value = os + 1
    End If
    UserForm1.ListBox4.RowSource = ""
    UserForm4.ListBox4.RowSource = ""
    tabela.ListRows.Add
    'Call IntervaloDados1
    'Call IntervaloDados
    UserForm4.Label10.Caption = Descr
    
    
    MsgBox "Cadastrado com sucesso!"
    'Call LimparCampos

bloqueado = False
End Sub
Sub EdtarCaixa()

End Sub
Sub DelCaixa()

End Sub

Sub LimparListaCaixa()
    Range("K5:R325").Select
    Application.CutCopyMode = False
    Selection.ClearContents
End Sub
Sub IntervaloDados2()

Dim base As Range
Dim Intc As Range
Dim destino As Range

Planilha6.Activate
Set base = Planilha6.Range("A1").CurrentRegion
Set Intc = Planilha6.Range("k1:r2")
Set destino = Planilha6.Range("k4:r4")
base.AdvancedFilter xlFilterCopy, Intc, destino
UserForm4.ListBox4.RowSource = Planilha6.Range("K4").CurrentRegion.Address

End Sub
Sub BaseCaixa()

    Sheets("Itens").Select
    Range("Tabela3[[#Headers],[Index]]").Select
    ActiveCell.FormulaR1C1 = "Index"
    Range("Tabela3[[#Headers],[Index]]").Select
    Application.Goto Reference:="Tabela3"
    Selection.Copy
    Sheets("ItensCaixa ").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("I2").Select
End Sub

Sub FiltroCaixaLocal()
' FiltroCaixaLocal Macro
Range("Tabela4[#All]").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange _
        :=Range("K1:R2"), Unique:=False
End Sub
Sub LimparFiltroCaixa()
' LimparFiltroCaixa Macro
   ActiveSheet.ShowAllData
End Sub

