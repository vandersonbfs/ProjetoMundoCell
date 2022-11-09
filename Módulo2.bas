Attribute VB_Name = "M�dulo2"
Option Explicit

Sub InserirProd()
bloqueado = True

    Dim tabela As ListObject
    Dim n As Long
    Dim id As Long
    Dim os As Long
    Dim hj As Date
    
    
    Set tabela = Planilha4.ListObjects(1)
    n = tabela.Range.Rows.Count
    id = Range("IDprod").Value
    os = Range("OS").Value
    hj = Range("HJ").Value
    
    tabela.Range(n, 1).Value = id
    tabela.Range(n, 2).Value = UserForm2.cbbMarcas.Value
    tabela.Range(n, 3).Value = UserForm2.cbbcategoria.Value
     tabela.Range(n, 4).Value = UserForm2.cbbModelos.Value
    tabela.Range(n, 5).Value = UserForm2.cbbFornecedor.Value
    tabela.Range(n, 7).Value = UserForm2.txtQuant.Value
    tabela.Range(n, 6).Value = UserForm2.txtValorEnt.Value
    tabela.Range(n, 8).Value = UserForm2.txtValorVen.Value
    
    
    UserForm1.ListBox5.RowSource = ""
    UserForm2.ListBox2.RowSource = ""
    tabela.ListRows.Add
    Call Atulizar_ListboxProd
    MsgBox "Cadastrado com sucesso!"
    Call LimparCamposProd
    
    Range("IDprod").Value = id + 1
    'Range("OS").Value = os + 1

bloqueado = False
End Sub
Sub EditarProd()
bloqueado = True

    Dim tabela As ListObject
    Dim n As Long
    Dim l As Long
    Dim os As Long
    Dim hj As Date
    
    Set tabela = Planilha4.ListObjects(1)
    n = UserForm2.ListBox2.Value
    l = tabela.Range.Columns().Find(n, , , xlWhole).Row
    
        tabela.Range(l, 2).Value = UserForm2.cbbMarcas.Value
        
        tabela.Range(l, 3).Value = UserForm2.cbbcategoria.Value
        
        tabela.Range(l, 4).Value = UserForm2.cbbModelos.Value
        
        tabela.Range(l, 5).Value = UserForm2.cbbFornecedor.Value
        
        tabela.Range(l, 6).Value = UserForm2.txtValorEnt.Value
        
        tabela.Range(l, 7).Value = UserForm2.txtQuant.Value
        
        tabela.Range(l, 8).Value = UserForm2.txtValorVen.Value
    
    MsgBox "Cadastro atualizado com sucesso!"
    Call Atulizar_ListboxProd
    Call LimparCamposProd
bloqueado = False

End Sub
Sub Atulizar_ListboxProd()
bloqueado = True
    Dim tabela As ListObject
    Set tabela = Planilha4.ListObjects(1)
    UserForm2.ListBox2.RowSource = tabela.DataBodyRange.Address(, , , True)
bloqueado = False
End Sub
Sub Atulizar_ListboxProd1()
bloqueado = True
    Dim tabela As ListObject
    Set tabela = Planilha4.ListObjects(1)
    UserForm1.ListBox5.RowSource = tabela.DataBodyRange.Address(, , , True)
bloqueado = False
End Sub
Sub DeletarProd()
bloqueado = True

    Dim tabela As ListObject
    Dim n As Long
    Dim l As Long
    Dim os As Long
    Dim hj As Date
    
    Set tabela = Planilha4.ListObjects(1)
    n = UserForm2.ListBox2.Value
    l = tabela.Range.Columns().Find(n, , , xlWhole).Row
    
    
    UserForm2.ListBox2.RowSource = ""
    'tabela.ListRows(l).Delete ----> Listrows come�a a conta apartir linha 2
    tabela.Range.Rows(l).Delete
    
    Call Atulizar_ListboxProd
    MsgBox "O registro foi deletado com sucesso!"
    Call LimparCamposProd

bloqueado = False
End Sub
Sub LimparCamposProd()
bloqueado = True

    UserForm2.cbbMarcas.Value = ""
    UserForm2.cbbModelos.Value = ""
    UserForm2.cbbFornecedor.Value = ""
    UserForm2.txtQuant.Value = ""
    UserForm2.txtValorEnt.Value = ""
    UserForm2.txtValorVen.Value = ""

bloqueado = False

End Sub
Sub IntervaloProdutos()
Dim base As Range
Dim Intc As Range
Dim destino As Range
Planilha4.Activate

Set base = Planilha4.Range("A1").CurrentRegion
Set Intc = Planilha4.Range("j1:q2")
Set destino = Planilha4.Range("s1:z1")


base.AdvancedFilter xlFilterCopy, Intc, destino

UserForm2.ListBox2.RowSource = Planilha4.Range("s1").CurrentRegion.Address

End Sub
