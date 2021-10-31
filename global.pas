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
       Msgbox "Inicio do Arquivo ", 65, "Aviso "
       Dyclientes.MoveNext
    End if
    
    Rem Se voltou um registro, logo não é novo
    Novo Registro = False
    
    LeRegistro
End Sub



