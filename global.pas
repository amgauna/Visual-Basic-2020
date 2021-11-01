Global BDados As Database, Dyclienes As Dynaset
Global CmCodigo As Field, CmpNome As Field
Global CmpTel As Field

' Formulário Principal

Dim NovoRegistro As Integer

Sub AtribuiCampos()
    CmpCodigo.Value = txtCodigo
    CmpNome.Value = txtNome
    CmpTel.Value = txtTel
End Sub

Sub BtnAnterior.BOF_click()
    DyClientes.MovPrevious
    
    If DyClientes.BOF then
       Msgbox "Inicio do Arquivo ", 65, " Aviso "
       Dyclientes.MoveNext
    End if
    
    Rem Se voltou um registro, logo não é novo
    Novo Registro = False
    
    LeRegistro
End Sub

Sub BtnApaga_Click()
    if MsgBox("Conforma a Deleção do Registro", 65, " Urgente ") = 1 then
       Dyclientes.Delete
       Dyclientes.movePrevius
       
       Rem Se voltou um registro, não é novo
       NovoRegistro = False
     
       LeRegistro
   End if
End Sub   

Sub BtnFiltro_Click()
    Load Frmfiltro
    Frmfiltro.Show
    
    if Dyclientes.BOF then
       MdgBox "Nenhum registro atendeu a condição", 65, " Aviso "
       Set DyClientes = BDados.CreateDynaset("Clientes")
    end if 
    
    Dyclientes.MoveFirst
    LimpaCampos
    InicializaCampos
    NovoRegistro = False
    LeRegistro
End Sub  

Sub BtnGrava_click()
    if NovoRegistro then
    
       DyClientes.FindFirst "Código = " & txtCódigo.Text
       
       if Not DyClientes.NoMath then
          MagBox "Código Inexistente", 65, " Aviso "
          exit sub
       end if
    end if
    
    if NovoRegistro then
       DyClientes.AddNew
    else
       Dyclientes.Edit
    end if
    
    AtribuiCampos
    DyClientes.Update
    NovoRegistro = False
End Sub



       

