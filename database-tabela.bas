Dim BDados as Database
Dim TbEstruturaclientes as TableDef

Set BDados = OpenDatabase 
  ("c:\tecnico\bdados.mdb")
  
  TbdEstruturaclientes.Connect = "OBDC; DNS=Srv_Corporativo, Database=Dados; UID=Juano; PWD=ObjectWay"
  
  TbdEstruturaClientes.SourceTable = "Clientes"
  
  TbsEstruturaClientes.Name = "OBDC Clientes"
  
  TbsEstruturaclientes.Atributes = DB+AttachsAvepwd
  
  BDados.TableDefs.Append TbdEstruturaClientes
  
  
