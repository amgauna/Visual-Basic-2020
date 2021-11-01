Set rsPessoa = New ADODB.Recordset
  
  rsPessoa.Open "SELECT * FROM Pessoa", gDB, 1, 3, 1
  rsPessoa.AdNew
  rsPessoa("Nome") = txtNome
  rsPessoa("Naturalidade") = txtNaturalidade
  rsPessoa("Estado") = cmEstado.ItemData(cmEstado.ListIndex)
  rsPessoa("Nascimento") = dtpNascimento
  rsPessoa("Emissao") = dtpEmissao
  rsPessoa.Update
  
  
