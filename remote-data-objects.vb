' RDO - Remote Data Objects 
' Ele foi uma alternativa a DAO antes da ADO ser lançada e se consolidar e funcionou como um, 
' vamos dizer, tampão, visto que geralmente na migração do RDO para a ADO o impacto era menor. 
' Os objetos RDO eram um boa opção para acesso a dados remotos, especificamente a servidores 
' SQL visto que a DAO é uma tecnologia mais adequada para comunicação com bases ISAM (dbase, 
' Paradox, Access, etc.) Os objetos RDO suportam o uso de cursores e a utilização de stored 
' procedures e apresentam a seguinte correlação com a tecnologia DAO:
' DAO/Jet	RDO
' DBEngine	rdoEngine
' Error	rdoError
' Workspace	rdoEnvironment
' DataBase	rdoConnection
' TableDef	rodTable
' Recordset	rdoResultset
' Field	rdoColumn
' QueryDef	rdoQuery
' Parameter	rdoParameter
' Como podemos notar o RDO foi praticamente um precursor da tecnologia ADO.

' A seguir no evento Click do botão de comando Conectar inclua o seguinte código:

Private Sub btnconectar_Click()

On Error GoTo trataerro

Dim cadeiaConexao As String

cadeiaConexao = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & servidor & "; DATABASE=" & cboBancodeDados & " ;PWD=" & password & "; UID=" & usuario & ";OPTION=3"

Set db = New rdoConnection

db.Connect = cadeiaConexao
db.CursorDriver = rdUseServer
db.EstablishConnection
Exit Sub

trataerro:
MsgBox Err.Number & " " & Err.Description, vbCritical

End Sub

' O código efetua a conexão com a base de dados informada no combobox usando o objeto RDO.
' O evento db_Connect deve possui o seguinte código:

Private Sub db_Connect(ByVal ErrorOccurred As Boolean)

On Error GoTo trataerro

Dim tabela As rdoTable
Dim existeTabela As Boolean

existeTabela = False
trocaBotoes True

For Each tabela In db.rdoTables
    tabelas.AddItem tabela.Name
    existeTabela = True
Next

If Not existeTabela Then
    MsgBox "O banco de dados esta vazio."
    btndesconectar_Click
End If
Exit Sub

trataerro:
MsgBox Err.Number & " " & Err.Description, vbCritical

End Sub

' O código acima verifica se existem tabelas na base de dados e inclui o nome
' de cada tabela no respectivo ListBox.
' Ao clicar em uma tabela queremos exibir os dados da mesma no outro ListBox 
' para isto inclua o código abaixo no evento Click do ListBox tabelas:

Private Sub tabelas_Click()

On Error GoTo trataerro


Dim tabela As String
Dim consulta As New rdoQuery
Dim resultados As rdoResultset
Dim conteudo_linha As String
Dim coluna As rdoColumn

conteudo.Clear

tabela = tabelas.List(tabelas.ListIndex)

Set consulta.ActiveConnection = db

consulta.SQL = "SELECT * FROM " & tabela & " WHERE 1"
consulta.Execute

Set resultados = consulta.OpenResultset

While Not resultados.EOF

conteudo_linha = ""

For Each coluna In resultados.rdoColumns
       conteudo_linha = conteudo_linha & coluna.Name & "=" & resultados(coluna.Name) & "; "
Next

conteudo.AddItem conteudo_linha
resultados.MoveNext
Wend

resultados.Close
Set resultados = Nothing
Exit Sub

trataerro:
MsgBox Err.Number & " " & Err.Description, vbCritical

End Sub

' Usando uma consulta via objeto rdoQuery estou gerando um recordset
' com os dados obtidos e exibindo no ListBox:
' No evento Unload estou efetuando o fechamento do banco de dados
' conforme código a seguir:

Private Sub Form_Unload(Cancel As Integer)

   If btnconectar.Enabled = False Then
          db.Close
   End If

End Sub

