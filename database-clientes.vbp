Dim BDados As Database
Dim DyClientes as Dynaset
Dim StrAtributos As String

StrAtributos = "Descripton = Banco de Dados do Servidor SrvEmpresa" & Chr$(13)
StrAtributos = AtrAtributos & "Network=DBNMP3" & Chr$(13)
StrAtributos = StrAtributos & "OemtoAnsi=No" & Chr$(13)
StrAtributos = AtrAtributos & "Address=\\SrvEmpresa\Pipe\SQL\Query" & Chr(13)
StrAtributos = StrAtributos & "DataBase=BDados"

RegisterDatabase "BDadosSQL", "SQL Server", True, StrAtributos

Set BDados = OpenDataBase("BDadosSQL", False, False, "OBDC;")
  
  Set Dyclientes = BDados.CreateDynaset("Clientes")
    
    
