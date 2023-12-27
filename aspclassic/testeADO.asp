<%@ Language=VBScript %>
<% 
    ' Receber os dados do formulário
    Dim campo1, campo2
    campo1 = Request.Form("campo1")
    campo2 = Request.Form("campo2")

    ' Conectar ao banco de dados usando autenticação do Windows
    Dim conexao, consulta
    Set conexao = Server.CreateObject("ADODB.Connection")

    ' Substitua SEU_SERVIDOR_SQL, SUA_BASE_DE_DADOS pelo nome do seu servidor SQL e sua base de dados
    ' Integrated Security=SSPI; = vai usar a integração de login com o usuário windows logado
    conexao.Open "Provider=SQLOLEDB;Data Source=SEU_SERVIDOR_SQL;Initial Catalog=SUA_BASE_DE_DADOS;Integrated Security=SSPI;"

    ' Inserir dados no banco de dados usando parâmetros parametrizados
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conexao
    cmd.CommandText = "INSERT INTO SuaTabela (NomeCampo1, NomeCampo2) VALUES (?, ?)"
    cmd.CommandType = adCmdText

    ' Adicionar parâmetros
    cmd.Parameters.Append(cmd.CreateParameter("@Campo1", adVarChar, adParamInput, 255, campo1))
    cmd.Parameters.Append(cmd.CreateParameter("@Campo2", adVarChar, adParamInput, 255, campo2))

    ' Executar o comando
    cmd.Execute , , adCmdText

    ' Fechar a conexão
    conexao.Close
    Set conexao = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Exemplo de Inserção no SQL Server com Autenticação do Windows</title>
</head>
<body>
    <h1>Dados inseridos com sucesso!</h1>
</body>
</html>