' Formulário Pesquisa

Sub Form_Load()
  DyClientes.MoveFirst
  Do While Not DyClientes.EOF
    LstEmpresas.AddItem
    CmpCodigo.Value + " " + CmpNome.Value
    DyClientes.moveNext
  Loop
End Sub

Sub LstEmpresas.DblClick()
  Dim Pos As Integer
  Dim Criterio As String
  Pos = InStr(LstEmpresas.Text," ")
  Criterio = "Codigo =" & Left(LstEmpresas.Text,Pos)
  DyClientes.FindFirst Criterio
  Unload FrmPesquisas
End Sub


  
  
  
    
