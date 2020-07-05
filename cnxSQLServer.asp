
<%
    'Variavel com a conexao do database
	'locaweb
 	Session("DBQM")="DRIVER={MySQL ODBC 5.3 ANSI Driver};SERVER=dbconseguro.mysql.dbaas.com.br;PORT=3306;DATABASE=dbconseguro;USER=dbconseguro;PASSWORD=Eth@43jkh#6;OPTION=3;"

	'Meu
	'Session("DBQM")="DRIVER={MySQL ODBC 5.1 Driver};SERVER=127.0.0.1;PORT=3306;DATABASE=dbconseguro;USER=root;PASSWORD=;OPTION=3;"
 
    'abro conexão     
  	Set MyConn=Server.CreateObject("ADODB.Connection")
  	MyConn.CommandTimeout = 999999   	
  	MyConn.ConnectionTimeout = 999999   	

    'Conexão
	MyConn.Open Session("DBQM")
%>