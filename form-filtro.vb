' Formulario Filtro

Sub BtnOk_Click()
  Dim StrSQL As String
  StrSQL = "Select * from Clientes Where"
  
  Rem Os campos caracter, memo e date tem um valor de tipo maior que 7
  
  if BDados.TableDefs("Clientes").Fields
    (Cmbcampo.Text).type > 7 then
    StrSQL = StrSQL & CmdCampo.Text & CmbOperador.Text & """" & TxtCond.Text & """"
  Else
    StrSQL = StrSQL & CmbCampo.Text & CmbOperador.Text & txtCond.Text
  End if
  
  Set Dyclientes = BDados.Createdynaset(StrSQL)
    Unload FrmFiltro
End Sub
  
Sub Form.Load()
    Dim i as integer
    
    Rem Adiciona ao Combo Box de Campos o nome dos campos da tabela
    
    For i=0 to
    BDados.TableDefs("Clientes").Fields.Count - 1
    CmbCampo.AddItem
    BDados.TableDefs("Clientes").Fields(i).Name
  Next    
    
  CmbOperador.AddItem "="   
  CmbOperador.AddItem ">"   
  CmbOperador.AddItem ">="   
  CmbOperador.AddItem "<"   
  CmbOperador.AddItem "<="   
  CmbOperador.AddItem "<>"   
  CmbOperador.AddItem "LIKE"   
Endsub  


  
  
  
