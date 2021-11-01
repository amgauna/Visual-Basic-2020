' Operação Valida

Dim BDados as Database
Dim Dyclientes as Dynaset
Dim StrConexao as String

StrConexao = "OBDC; DSN=Srv_Corporativo; Database=BDados; UID=Juano; PWD="ObjectWay"

Set BDados = OpenDatabase("""", False, False, SRTConexao)
  
  Set Dyclienes = BDados.CreateDynaset("Clientes")
    
    
