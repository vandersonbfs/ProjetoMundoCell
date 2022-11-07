Attribute VB_Name = "Módulo3"
Sub InserirItem()
bloqueado = True

    Dim tabela As ListObject
    Dim n As Integer
    Dim id As Integer
    Dim os As Long
    Dim hj As Date
    Dim Qntd As Integer
    Dim vlr As Long
    
    
    Set tabela = Planilha3.ListObjects(1)
    n = tabela.Range.Rows.Count
    id = Range("IDitem").Value
    os = Range("Slv").Value
    hj = Range("HJ").Value
    
    tabela.Range(n, 1).Value = id
    tabela.Range(n, 2).Value = os
    tabela.Range(n, 3).Value = UserForm3.cbbCatItem.Value
    tabela.Range(n, 4).Value = UserForm3.ccbMarcaItem.Value
    
    tabela.Range(n, 5).Value = UserForm3.cbbItemItem.Value
    Qntd = UserForm3.txtQuantItem.Value
    tabela.Range(n, 6).Value = Qntd
    tabela.Range(n, 7).Value = UserForm3.txtValItem.Value
    
    
    Range("IDitem").Value = id + 1
    'Range("OS").Value = os + 1
    UserForm1.ListBox6.RowSource = ""
    UserForm3.ListBox3.RowSource = ""
   
    
    tabela.ListRows.Add
    Call IntervaloDados1
    Call IntervaloDados
    Call BaseCaixa
    MsgBox "Cadastrado com sucesso!"
    'Call LimparCampos
    bloqueado = False
End Sub
Sub EditarItem()
bloqueado = True

    Dim tabela As ListObject
    Dim n As Integer
    Dim l As Integer
    Dim os As Long
    Dim hj As Date
    
    Set tabela = Planilha3.ListObjects(1)
    n = UserForm3.ListBox3.Value
    l = tabela.Range.Columns().Find(n, , , xlWhole).Row
    
    tabela.Range(l, 3).Value = UserForm3.cbbCatItem.Value
    tabela.Range(l, 4).Value = UserForm3.ccbMarcaItem.Value
    tabela.Range(l, 5).Value = UserForm3.cbbItemItem.Value
    tabela.Range(l, 6).Value = UserForm3.txtQuantItem.Value
    tabela.Range(l, 7).Value = UserForm3.txtValItem.Value
    
    
    MsgBox "Cadastro atualizado com sucesso!"
    'Call Atulizar_ListBox3
    Call LimparCampos
    bloqueado = False
    Call BaseCaixa

End Sub
Sub Atulizar_Listbox3()
bloqueado = True
    Dim tabela As ListObject
    Set tabela = Planilha3.ListObjects(1)
    UserForm3.ListBox3.RowSource = tabela.DataBodyRange.Address(, , , True)
bloqueado = False
End Sub
Sub DeletarItem()
bloqueado = True

    Dim tabela As ListObject
    Dim n As Integer
    Dim l As Integer
    Dim os As Long
    Dim hj As Date
    
    Set tabela = Planilha3.ListObjects(1)
    n = UserForm3.ListBox3.Value
    l = tabela.Range.Columns().Find(n, , , xlWhole).Row
    
    UserForm3.ListBox3.RowSource = ""
    'tabela.ListRows(l).Delete ----> Listrows começa a conta apartir linha 2
    tabela.Range.Rows(l).Delete
    
    Call Atulizar_Listbox3
    MsgBox "O registro foi deletado com sucesso!"
    Call LimparCampos
    Call BaseCaixa
bloqueado = False
End Sub
Sub LimparCamposItens()
bloqueado = True

    UserForm3.cbbCatItem.Value = ""
    UserForm3.ccbMarcaItem.Value = ""
    UserForm3.cbbItemItem.Value = ""
    UserForm3.txtQuantItem.Value = ""
    UserForm3.txtValItem.Value = ""

bloqueado = False

End Sub
Sub IntervaloDados1()
Dim base As Range
Dim Intc As Range
Dim destino As Range
Planilha3.Activate

Set base = Planilha3.Range("A1").CurrentRegion
Set Intc = Planilha3.Range("k1:q2")
Set destino = Planilha3.Range("k4:q4")

base.AdvancedFilter xlFilterCopy, Intc, destino

UserForm3.ListBox3.RowSource = Planilha3.Range("K4").CurrentRegion.Address

End Sub
Sub Atulizar_ListboxProdAdd()
bloqueado = True
    Dim tabela As ListObject
    Set tabela = Planilha4.ListObjects(1)
    UserForm3.ListBox3.RowSource = tabela.DataBodyRange.Address(, , , True)
bloqueado = False
End Sub
