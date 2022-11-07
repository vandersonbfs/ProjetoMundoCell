Attribute VB_Name = "M�dulo1"
Option Explicit
Global bloqueado As Boolean
'Metodos para inserção e manipulação dos dados
Sub Inserir()
bloqueado = True
    
    Dim tabela As ListObject
    Dim n As Integer
    Dim id As Integer
    Dim os As Long
    Dim hj As Date
   
    
    
    Set tabela = Planilha1.ListObjects(1)
    n = tabela.Range.Rows.Count
    id = Range("ID").Value
    os = Range("OS").Value
    hj = Range("HJ").Value
    idclient = Range("IDClient").Value
    
    
    tabela.Range(n, 1).Value = id
    tabela.Range(n, 2).Value = os
    tabela.Range(n, 3).Value = UserForm1.txtnome.Value
    tabela.Range(n, 4).Value = UserForm1.txttelefone.Value
    tabela.Range(n, 5).Value = UserForm1.cbbMarca.Value
    tabela.Range(n, 6).Value = hj
    tabela.Range(n, 7).Value = UserForm1.cbbModelo.Value
    tabela.Range(n, 8).Value = UserForm1.cbbServi�o.Value
    tabela.Range(n, 9).Value = UserForm1.cbbFormaPagamento.Value
    tabela.Range(n, 10).Value = UserForm1.cbbRecebido.Value
    tabela.Range(n, 11).Value = UserForm1.txtcpf.Value
    tabela.Range(n, 12).Value = UserForm1.txtOBS.Value
    tabela.Range(n, 13).Value = UserForm1.CheckBox1.Value
    tabela.Range(n, 14).Value = UserForm1.CheckBox2.Value
    tabela.Range(n, 15).Value = UserForm1.CheckBox3.Value
    tabela.Range(n, 16).Value = UserForm1.CheckBox4.Value
    tabela.Range(n, 17).Value = UserForm1.CheckBox5.Value
    tabela.Range(n, 18).Value = UserForm1.CheckBox6.Value
    tabela.Range(n, 19).Value = UserForm1.CheckBox7.Value
    tabela.Range(n, 20).Value = UserForm1.CheckBox8.Value
    tabela.Range(n, 21).Value = UserForm1.CheckBox9.Value
    tabela.Range(n, 22).Value = UserForm1.txtEmail.Value
    tabela.Range(n, 23).Value = UserForm1.cbbStatus.Value
    tabela.Range(n, 24).Value = UserForm1.txtSerie.Value
    
    Range("ID").Value = id + 1
    Range("OS").Value = os + 1
    Range("IDClient").Value = idclient + 1
    
    
    
    
    
    UserForm1.ListBox1.RowSource = ""
    tabela.ListRows.Add
    
        Set tabela = Planilha9.ListObjects(1)
        n = tabela.Range.Rows.Count
        id = Range("IDClient").Value
        
        
        tabela.Range(n, 1).Value = id
        tabela.Range(n, 2).Value = os
        tabela.Range(n, 2).Value = UserForm1.txtnome.Value
        tabela.Range(n, 3).Value = UserForm1.txttelefone.Value
        tabela.Range(n, 4).Value = UserForm1.txtEmail.Value
        tabela.Range(n, 5).Value = UserForm1.txtcpf.Value
        tabela.ListRows.Add
        
    Call Ordenar
    Call Atulizar_Listbox1
    Call LimparCampos
    MsgBox "Cadastrado com sucesso!"
    
bloqueado = False
End Sub
Sub Editar()
bloqueado = True
    
    Dim tabela As ListObject
    Dim n As Integer
    Dim l As Integer
    Dim os As Long
    Dim hj As Date
    Dim id As Long
    Dim slv As Long
    
    
    Set tabela = Planilha1.ListObjects(1)
    n = UserForm1.ListBox1.Value
    l = tabela.Range.Columns().Find(n, , , xlWhole).Row
    
    tabela.Range(l, 3).Value = UserForm1.txtnome.Value
    tabela.Range(l, 4).Value = UserForm1.txttelefone.Value
    tabela.Range(l, 5).Value = UserForm1.cbbMarca.Value
    'tabela.Range(l, 6).Value = hj
    tabela.Range(l, 7).Value = UserForm1.cbbModelo.Value
    tabela.Range(l, 8).Value = UserForm1.cbbServi�o.Value
    tabela.Range(l, 9).Value = UserForm1.cbbFormaPagamento.Value
    tabela.Range(l, 10).Value = UserForm1.cbbRecebido.Value
    tabela.Range(l, 11).Value = UserForm1.txtcpf.Value
    tabela.Range(l, 12).Value = UserForm1.txtOBS.Value
    tabela.Range(l, 13).Value = UserForm1.CheckBox1.Value
    tabela.Range(l, 14).Value = UserForm1.CheckBox2.Value
    tabela.Range(l, 15).Value = UserForm1.CheckBox3.Value
    tabela.Range(l, 16).Value = UserForm1.CheckBox4.Value
    tabela.Range(l, 17).Value = UserForm1.CheckBox5.Value
    tabela.Range(l, 18).Value = UserForm1.CheckBox6.Value
    tabela.Range(l, 19).Value = UserForm1.CheckBox7.Value
    tabela.Range(l, 20).Value = UserForm1.CheckBox8.Value
    tabela.Range(l, 21).Value = UserForm1.CheckBox9.Value
    tabela.Range(l, 22).Value = UserForm1.txtEmail.Value
    tabela.Range(l, 23).Value = UserForm1.cbbStatus.Value
    tabela.Range(l, 24).Value = UserForm1.txtSerie.Value
    

    Call Atulizar_Listbox1
    Call LimparCampos
bloqueado = False


End Sub
Sub Atulizar_Listbox1()
bloqueado = True
    Dim tabela As ListObject
    Set tabela = Planilha1.ListObjects(1)
    UserForm1.ListBox1.RowSource = tabela.DataBodyRange.Address(, , , True)
bloqueado = False
End Sub
Sub Atulizar_ListBox4()
bloqueado = True
    Dim tabela As ListObject
    Set tabela = Planilha5.ListObjects(1)
    UserForm1.ListBox4.RowSource = tabela.DataBodyRange.Address(, , , True)
bloqueado = False
End Sub
Sub Deletar()
bloqueado = True

    Dim tabela As ListObject
    Dim n As Integer
    Dim l As Integer
    Dim os As Long
    Dim hj As Date
    
    Set tabela = Planilha1.ListObjects(1)
    n = UserForm1.ListBox1.Value
    l = tabela.Range.Columns().Find(n, , , xlWhole).Row
    
    UserForm1.ListBox1.RowSource = ""
    'tabela.ListRows(l).Delete ----> Listrows come�a a conta apartir linha 2
    tabela.Range.Rows(l).Delete
    
    Call Atulizar_Listbox1
    MsgBox "O registro foi deletado com sucesso!"
    Call LimparCampos
bloqueado = False
End Sub
Sub LimparCampos()
bloqueado = True

    UserForm1.txtnome.Value = ""
    UserForm1.txttelefone.Value = ""
    UserForm1.cbbMarca.Value = ""
    UserForm1.cbbModelo.Value = ""
    UserForm1.cbbServi�o.Value = ""
    UserForm1.cbbFormaPagamento.Value = ""
    UserForm1.cbbRecebido.Value = ""
    UserForm1.txtcpf.Value = ""
    UserForm1.txtOBS.Value = ""
    UserForm1.CheckBox1.Value = False
    UserForm1.CheckBox2.Value = False
    UserForm1.CheckBox3.Value = False
    UserForm1.CheckBox4.Value = False
    UserForm1.CheckBox5.Value = False
    UserForm1.CheckBox6.Value = False
    UserForm1.CheckBox7.Value = False
    UserForm1.CheckBox8.Value = False
    UserForm1.CheckBox9.Value = False
    UserForm1.txtEmail.Value = ""
    UserForm1.cbbStatus.Value = ""
    Planilha3.Range("k2", "Q2").Clear
    Planilha6.Range("k2", "R2").Clear
    Planilha3.Range("k5", "R30").Clear
    Planilha6.Range("k5", "R30").Clear
    Planilha6.Range("U2").Clear
    Planilha6.Range("U4").Clear
    Planilha4.Range("j2", "Q2").Clear
    
    
bloqueado = False

End Sub
Sub IntervaloDados()
Dim base As Range
Dim Intc As Range
Dim destino As Range
Planilha3.Activate

Set base = Planilha3.Range("A1").CurrentRegion
Set Intc = Planilha3.Range("k1:q2")
Set destino = Planilha3.Range("k4:q4")

base.AdvancedFilter xlFilterCopy, Intc, destino

UserForm1.ListBox6.RowSource = Planilha3.Range("K4").CurrentRegion.Address
UserForm1.Label63.Caption = Planilha3.Range("L2")
UserForm1.Label64.Caption = Planilha3.Range("T1")
UserForm1.Label66.Caption = Planilha3.Range("T2")
UserForm1.Label68.Caption = Planilha3.Range("T3")
End Sub

Sub testeList()

UserForm1.ListBox6.RowSource = Planilha3.Range("H3").CurrentRegion.Address

End Sub
Sub Ordenar()

    Application.Goto Reference:="Tabela1"
    ActiveWorkbook.Worksheets("Base").ListObjects("Tabela1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Base").ListObjects("Tabela1").Sort.SortFields.Add _
        Key:=Range("Tabela1[OS]"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Base").ListObjects("Tabela1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub EnviarEmail()

On Error Resume Next

If UserForm1.txtEmail.Value = "" Then
MsgBox "Favor preencher o campo Email!"
Else
    Call Pdf
    Set objeto_outlook = CreateObject("Outlook.Application")
    
    Set Email = objeto_outlook.createitem(0)
    
    os = Planilha2.Range("Slv")
    
    Email.display
    
    Email.To = txtEmail.Value
    
    Email.Subject = "Mundo CELL - Ordem de Servi�o N� " & os
    
    texto = Planilha2.Range("F3").Value
    assinatura = Planilha2.Range("F2").Value

    Email.Body = "Ol� " & txtnome.Value & "," & Chr(10) & Chr(10) _
    & texto & Chr(10) & Chr(10) _
    & "Atenciosamente," & Chr(10) & assinatura
    
    Email.Attachments.Add (ThisWorkbook.Path & "\OS_" & os & ".pdf")
    Email.send
End If

End Sub
Sub FinalizarServico()

cbbStatus.value =

End Sub
