VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Conectiva Sistemas"
   ClientHeight    =   9450.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14670
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btDeletar_Click()
Dim nlin As Integer

    If TbgEditar.Value = True Then
        nlin = ListBox1.ListIndex
        
        If nlin = -1 Then
            MsgBox "Selecione um item para ser deletado!"
            Exit Sub
        ElseIf ListBox1.Value = 0 Then
        
        MsgBox "Selecione um item para ser deletado!"
        Exit Sub
        End If
        Call Deletar
        Else
        MsgBox "Clique em Editar primeiro!"
    End If
End Sub

Private Sub btNovo_Click()
Call LimparCampos
End Sub


Private Sub btPesquisar_Click()
UserForm2.Show
End Sub

Private Sub btSalvar_Click()
Dim nlin As Integer

    If TbgEditar.Value = True Then
        nlin = ListBox1.ListIndex
        
        If nlin = -1 Then
            MsgBox "Selecione um item para ser editado!"
            Exit Sub
        ElseIf ListBox1.Value = 0 Then
        
        MsgBox "Selecione um item para ser editado!"
        Exit Sub
        End If
        Call Editar
    Else
        Call Inserir
    End If
End Sub



Private Sub cbbEmail_Click()
On Error Resume Next

If UserForm1.txtEmail.Value = "" Then
MsgBox "Favor preencher o campo Email!"
Else
    Call Pdf
    Set objeto_outlook = CreateObject("Outlook.Application")
    
    Set Email = objeto_outlook.createitem(0)
    
    os = Planilha2.Range("Slv")
    'Empresa
    emp = Planilha2.Range("Empresa")
    
    Email.display
    
    Email.To = txtEmail.Value
    
    Email.Subject = "Ordem de Serviço Nº " & os & " - " & emp
    
    texto = Planilha2.Range("F3").Value
    assinatura = Planilha2.Range("F2").Value

    Email.Body = "Olá " & txtnome.Value & "," & Chr(10) & Chr(10) _
    & texto & Chr(10) & Chr(10) _
    & "Atenciosamente," & Chr(10) & assinatura
    
    Email.Attachments.Add (ThisWorkbook.Path & "\OS_" & os & ".pdf")
    Email.send
End If
End Sub

Private Sub cbbMarca_Change()
 If cbbMarca.Value = "Apple" Then
        Dim totLIN As Long
        Dim LIN As Long
        UserForm1.cbbModelo.Clear
        totLIN = Planilha2.Range("S" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbModelo.AddItem Planilha2.Cells(LIN, 19)
        Next LIN
 ElseIf cbbMarca.Value = "Asus" Then
        UserForm1.cbbModelo.Clear
        totLIN = Planilha2.Range("t" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbModelo.AddItem Planilha2.Cells(LIN, 20)
        Next LIN
 ElseIf cbbMarca.Value = "Google" Then
        UserForm1.cbbModelo.Clear
        totLIN = Planilha2.Range("u" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbModelo.AddItem Planilha2.Cells(LIN, 21)
        Next LIN
 ElseIf cbbMarca.Value = "Huawei" Then
        UserForm1.cbbModelo.Clear
        totLIN = Planilha2.Range("v" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbModelo.AddItem Planilha2.Cells(LIN, 22)
        Next LIN
   ElseIf cbbMarca.Value = "LG" Then
        UserForm1.cbbModelo.Clear
        totLIN = Planilha2.Range("x" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbModelo.AddItem Planilha2.Cells(LIN, 24)
        Next LIN
  ElseIf cbbMarca.Value = "Motorola" Then
        UserForm1.cbbModelo.Clear
        totLIN = Planilha2.Range("w" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbModelo.AddItem Planilha2.Cells(LIN, 23)
        Next LIN
  ElseIf cbbMarca.Value = "Multilaser" Then
        UserForm1.cbbModelo.Clear
        totLIN = Planilha2.Range("y" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbModelo.AddItem Planilha2.Cells(LIN, 25)
        Next LIN
  ElseIf cbbMarca.Value = "Nokia" Then
        UserForm1.cbbModelo.Clear
        totLIN = Planilha2.Range("z" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbModelo.AddItem Planilha2.Cells(LIN, 26)
        Next LIN
   ElseIf cbbMarca.Value = "OBABOX" Then
        UserForm1.cbbModelo.Clear
        totLIN = Planilha2.Range("aa" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbModelo.AddItem Planilha2.Cells(LIN, 27)
        Next LIN
  ElseIf cbbMarca.Value = "OnePlus" Then
        UserForm1.cbbModelo.Clear
        totLIN = Planilha2.Range("ab" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbModelo.AddItem Planilha2.Cells(LIN, 28)
        Next LIN
  ElseIf cbbMarca.Value = "Positivo" Then
        UserForm1.cbbModelo.Clear
        totLIN = Planilha2.Range("ac" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbModelo.AddItem Planilha2.Cells(LIN, 29)
        Next LIN
  ElseIf cbbMarca.Value = "Realme" Then
        UserForm1.cbbModelo.Clear
        totLIN = Planilha2.Range("ad" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbModelo.AddItem Planilha2.Cells(LIN, 30)
        Next LIN
  ElseIf cbbMarca.Value = "Samsung" Then
        UserForm1.cbbModelo.Clear
        totLIN = Planilha2.Range("ae" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbModelo.AddItem Planilha2.Cells(LIN, 31)
        Next LIN
  ElseIf cbbMarca.Value = "Sony" Then
        UserForm1.cbbModelo.Clear
        totLIN = Planilha2.Range("af" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbModelo.AddItem Planilha2.Cells(LIN, 32)
        Next LIN
  ElseIf cbbMarca.Value = "Xiaomi" Then
        UserForm1.cbbModelo.Clear
        totLIN = Planilha2.Range("ag" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbModelo.AddItem Planilha2.Cells(LIN, 33)
        Next LIN
  ElseIf cbbMarca.Value = "Alcatel" Then
        UserForm1.cbbModelo.Clear
        totLIN = Planilha2.Range("ah" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbModelo.AddItem Planilha2.Cells(LIN, 34)
        Next LIN
  ElseIf cbbMarca.Value = "Huawei" Then
        UserForm1.cbbModelo.Clear
        totLIN = Planilha2.Range("ai" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbModelo.AddItem Planilha2.Cells(LIN, 35)
        Next LIN
  ElseIf cbbMarca.Value = "HTC" Then
        UserForm1.cbbModelo.Clear
        totLIN = Planilha2.Range("aj" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbModelo.AddItem Planilha2.Cells(LIN, 36)
        Next LIN
  ElseIf cbbMarca.Value = "Lenovo" Then
        UserForm1.cbbModelo.Clear
        totLIN = Planilha2.Range("aj" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbModelo.AddItem Planilha2.Cells(LIN, 37)
        Next LIN
 Else
    'MsgBox "Escolha a marca!"
 End If
End Sub

Private Sub CommandButton15_Click()

End Sub

Private Sub cbbStatus_Change()

End Sub

Private Sub CommandButton5_Click()
UserForm4.Show
End Sub

Private Sub CommandButton7_Click()
UserForm3.Show
End Sub



Private Sub ListBox1_Change()
Dim nlin As Integer
Dim slv As Long


Call LimpaFiltro
nlin = ListBox1.ListIndex
If bloqueado = True Then Exit Sub
If nlin = -1 Then Exit Sub
txtnome.Value = ListBox1.List(nlin, 2)
txttelefone.Value = ListBox1.List(nlin, 3)
cbbMarca.Value = ListBox1.List(nlin, 4)
cbbModelo.Value = ListBox1.List(nlin, 6)
cbbServiço.Value = ListBox1.List(nlin, 7)
cbbFormaPagamento.Value = ListBox1.List(nlin, 8)
cbbRecebido.Value = ListBox1.List(nlin, 9)
txtcpf.Value = ListBox1.List(nlin, 10)
txtOBS.Value = ListBox1.List(nlin, 11)
CheckBox1.Value = ListBox1.List(nlin, 12)
CheckBox2.Value = ListBox1.List(nlin, 13)
CheckBox3.Value = ListBox1.List(nlin, 14)
CheckBox4.Value = ListBox1.List(nlin, 15)
CheckBox5.Value = ListBox1.List(nlin, 16)
CheckBox6.Value = ListBox1.List(nlin, 17)
CheckBox7.Value = ListBox1.List(nlin, 18)
CheckBox8.Value = ListBox1.List(nlin, 19)
CheckBox9.Value = ListBox1.List(nlin, 20)
txtEmail.Value = ListBox1.List(nlin, 21)
cbbStatus.Value = ListBox1.List(nlin, 22)
txtSerie = ListBox1.List(nlin, 23)

Planilha3.Range("L2").Value = ListBox1.List(nlin, 1)
Planilha9.Range("M2").Value = UserForm1.txtnome.Value

slv = ListBox1.List(nlin, 1)
Range("Slv").Value = slv
Planilha6.Range("L2").Value = slv
Planilha6.Range("L2").Value = ListBox1.List(nlin, 1)

Call IntervaloDados

Label63.Caption = Planilha6.Range("L2")     'Numero da OS
Label64.Caption = Planilha6.Range("U1")     'Valor total
Label66.Caption = Planilha6.Range("U11")    'Itens
Label68.Caption = Planilha6.Range("U10")    'Quantidade

Planilha7.Range("H2").Value = ListBox1.List(nlin, 1)
Planilha7.Range("H3").Value = ListBox1.List(nlin, 5)
'Finalização    H4 - Criar evento de finalização de OS
Planilha7.Range("H5").Value = ListBox1.List(nlin, 22)
Planilha7.Range("C8").Value = ListBox1.List(nlin, 2)
Planilha7.Range("C9").Value = ListBox1.List(nlin, 3)
Planilha7.Range("F9").Value = ListBox1.List(nlin, 21)
Planilha7.Range("C12").Value = ListBox1.List(nlin, 4)
Planilha7.Range("C13").Value = ListBox1.List(nlin, 6)
Planilha7.Range("G8").Value = ListBox1.List(nlin, 10)
Planilha7.Range("D13").Value = ListBox1.List(nlin, 6)
Planilha7.Range("D14").Value = ListBox1.List(nlin, 23)
Planilha7.Range("F13").Value = ListBox1.List(nlin, 11)
Call IntervaloDados2
If Planilha6.Range("U4").Value <> "" Then
    Planilha7.Range("H29").Value = Planilha6.Range("U4").Value
    Planilha7.Range("H30").Value = Planilha6.Range("U5").Value
Else
    Planilha7.Range("H29").Value = Planilha6.Range("U4").Value
    Planilha7.Range("H30").Value = Planilha6.Range("U1").Value
End If

End Sub



Private Sub ListBox5_Click()
Dim nlin As Integer
    nlin = UserForm1.ListBox5.ListIndex
        If bloqueado = True Then Exit Sub
        If nlin = -1 Then Exit Sub
        
    tstMarcaprodt.Value = ListBox5.List(nlin, 1)
    txtCatPrdo.Value = ListBox5.List(nlin, 2)
    txtModeloProdt.Value = ListBox5.List(nlin, 3)
    txtFornProdt.Value = ListBox5.List(nlin, 4)
    txtValCompProdt.Value = ListBox5.List(nlin, 5)
    txtQntdProdt.Value = ListBox5.List(nlin, 6)
    txtValVendProdt.Value = ListBox5.List(nlin, 7)
End Sub

Private Sub MultiPage1_Change()
Call Atulizar_ListboxProd1
End Sub

Private Sub TbgEditar_Click()

End Sub

Private Sub UserForm_Activate()
    Dim totLIN As Long
    Dim LIN As Long
    Planilha6.Range("L2").Value = ""
    
    UserForm1.cbbMarca.Clear
    totLIN = Planilha2.Range("h" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbMarca.AddItem Planilha2.Cells(LIN, 8)
        Next LIN

End Sub

Private Sub UserForm_Initialize()
'Remover o comentario abaixo
'Application.Visible = False
MultiPage1.BackColor = &HFFFFFF



Call Atulizar_Listbox1
Call Atulizar_ListBox4
Call Atulizar_ListboxProd1
Call BaseCaixa
Call LimpaFiltro
'Planilha5.Range("L2").Value = ""
Label56.Caption = Planilha5.Range("L4").Value
Label58.Caption = Planilha5.Range("L5").Value
Label61.Caption = Planilha5.Range("L6").Value

    Dim totLIN As Long
    Dim LIN As Long
    UserForm1.cbbFormaPagamento.Clear
    totLIN = Planilha2.Range("k" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbFormaPagamento.AddItem Planilha2.Cells(LIN, 11)
        Next LIN
        
    UserForm1.cbbServiço.Clear
    totLIN = Planilha2.Range("g" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbServiço.AddItem Planilha2.Cells(LIN, 7)
        Next LIN
        
    UserForm1.cbbRecebido.Clear
    totLIN = Planilha2.Range("L" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbRecebido.AddItem Planilha2.Cells(LIN, 12)
        Next LIN
                
    UserForm1.cbbStatus.Clear
    totLIN = Planilha2.Range("L" & Rows.Count).End(xlUp).Row
    
        For LIN = 1 To totLIN
            UserForm1.cbbStatus.AddItem Planilha2.Cells(LIN, 13)
        Next LIN
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Application.Visible = True
End Sub
