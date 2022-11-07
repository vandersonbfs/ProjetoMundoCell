Attribute VB_Name = "M�dulo5"
Sub Pdf()
    os = Planilha2.Range("Slv")
    Planilha7.Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
    "E+:\Users\Vanderson\Documents\OS_" & os & ".pdf", Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
    True
End Sub

Sub InserirProdCaixa()
bloqueado = True

    Dim tabela As ListObject
    Dim n As Integer
    Dim id As Integer
    Dim os As Long
    Dim hj As Date
    
    
    Set tabela = Planilha4.ListObjects(1)
    n = tabela.Range.Rows.Count
    id = Range("IDprod").Value
    os = Range("OS").Value
    hj = Range("HJ").Value
    
    tabela.Range(n, 1).Value = id
    tabela.Range(n, 2).Value = UserForm5.cbbMarcas.Value
    tabela.Range(n, 3).Value = UserForm5.cbbcategoria.Value
     tabela.Range(n, 4).Value = UserForm5.cbbModelos.Value
    tabela.Range(n, 5).Value = UserForm5.cbbFornecedor.Value
    tabela.Range(n, 7).Value = UserForm5.txtQuant.Value
    tabela.Range(n, 6).Value = UserForm5.txtValorEnt.Value
    tabela.Range(n, 8).Value = UserForm5.txtValorVen.Value
    
    
    UserForm1.ListBox5.RowSource = ""
    UserForm5.ListBox2.RowSource = ""
    tabela.ListRows.Add
    Call Atulizar_ListboxProd
    MsgBox "Cadastrado com sucesso!"
    Call LimparCamposProd
    
    Range("IDprod").Value = id + 1
    'Range("OS").Value = os + 1

bloqueado = False
End Sub
Sub EditarProdCaixa()
bloqueado = True

    Dim tabela As ListObject
    Dim n As Integer
    Dim l As Integer
    Dim os As Long
    Dim hj As Date
    
    Set tabela = Planilha4.ListObjects(1)
    n = UserForm5.ListBox2.Value
    l = tabela.Range.Columns().Find(n, , , xlWhole).Row
    
        tabela.Range(l, 2).Value = UserForm5.cbbMarcas.Value
        
        tabela.Range(l, 3).Value = UserForm5.cbbcategoria.Value
        
        tabela.Range(l, 4).Value = UserForm5.cbbModelos.Value
        
        tabela.Range(l, 5).Value = UserForm5.cbbFornecedor.Value
        
        tabela.Range(l, 6).Value = UserForm5.txtValorEnt.Value
        
        tabela.Range(l, 7).Value = UserForm5.txtQuant.Value
        
        tabela.Range(l, 8).Value = UserForm5.txtValorVen.Value
    
    MsgBox "Cadastro atualizado com sucesso!"
    Call Atulizar_ListboxProd
    Call LimparCamposProd
bloqueado = False

End Sub
Sub Atulizar_ListboxProdCaixa()
bloqueado = True
    Dim tabela As ListObject
    Set tabela = Planilha4.ListObjects(1)
    UserForm5.ListBox2.RowSource = tabela.DataBodyRange.Address(, , , True)
bloqueado = False
End Sub
Sub Atulizar_ListboxProd1Caixa()
bloqueado = True
    Dim tabela As ListObject
    Set tabela = Planilha4.ListObjects(1)
    UserForm1.ListBox5.RowSource = tabela.DataBodyRange.Address(, , , True)
bloqueado = False
End Sub
Sub DeletarProdCaixa()
bloqueado = True

    Dim tabela As ListObject
    Dim n As Integer
    Dim l As Integer
    Dim os As Long
    Dim hj As Date
    
    Set tabela = Planilha4.ListObjects(1)
    n = UserForm5.ListBox2.Value
    l = tabela.Range.Columns().Find(n, , , xlWhole).Row
    
    
    UserForm5.ListBox2.RowSource = ""
    'tabela.ListRows(l).Delete ----> Listrows come�a a conta apartir linha 2
    tabela.Range.Rows(l).Delete
    
    Call Atulizar_ListboxProd
    MsgBox "O registro foi deletado com sucesso!"
    Call LimparCamposProd

bloqueado = False
End Sub
Sub LimparCamposProdCaixa()
bloqueado = True

    UserForm5.cbbMarcas.Value = ""
    UserForm5.cbbModelos.Value = ""
    UserForm5.cbbFornecedor.Value = ""
    UserForm5.txtQuant.Value = ""
    UserForm5.txtValorEnt.Value = ""
    UserForm5.txtValorVen.Value = ""

bloqueado = False

End Sub
Sub IntervaloProdutosCaixa()
Dim base As Range
Dim Intc As Range
Dim destino As Range
Planilha4.Activate

Set base = Planilha4.Range("A1").CurrentRegion
Set Intc = Planilha4.Range("j1:q2")
Set destino = Planilha4.Range("s1:z1")


base.AdvancedFilter xlFilterCopy, Intc, destino

UserForm5.ListBox2.RowSource = Planilha4.Range("s1").CurrentRegion.Address

End Sub
Sub LimpaFiltro()
'
' LimpaFiltro Macro
'

'
    Columns("A:H").Select
    ActiveSheet.ShowAllData
End Sub
Sub ConverterNumeros()
'
' ConverterNumeros Macro
'
    Range("Q5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.ScrollRow = 1048525
    ActiveWindow.ScrollRow = 1038673
    ActiveWindow.ScrollRow = 1017561
    ActiveWindow.ScrollRow = 327928
    ActiveWindow.ScrollRow = 263187
    ActiveWindow.ScrollRow = 156224
    ActiveWindow.ScrollRow = 73186
    ActiveWindow.ScrollRow = 46445
    ActiveWindow.ScrollRow = 30964
    ActiveWindow.ScrollRow = 8445
    ActiveWindow.ScrollRow = 1
    Selection.TextToColumns Destination:=Range("Q5"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
End Sub

