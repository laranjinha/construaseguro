<!--#include file="funcoes.asp"-->
<%
	'Zero variaveis de sessão
   	Session("gvMenu")="0"
%>
<!DOCTYPE html>
<html lang="pt-br">

	<head>
        <!--#include file="head.asp"-->
	</head>

	<body  onload="horizontal();">

        <!--#include file="cabecalho.asp"-->

		<main class="principal">
			    <div id="caption" align="center">
					<p>&nbsp;</p>
					<h3>Problemas</h3>
					<p>&nbsp;</p>
					<p>Foi detectado que seu browser est&aacute; configurado para n&atilde;o executar scripts!</p>
					<p>&Eacute; necess&aacute;rio habilitar a execu&ccedil;&atilde;o de scripts para execu&ccedil;&atilde;o deste site.</p>
					<p>Acerte a configura&ccedil;&atilde;o do seu browser e tente se logar novamente.</p>
					<p>&nbsp;</p>
					<h3><a href="logout.asp">Tentar novamente</a></h3>
				</div>	
        </main>

		<!--#include file="rodape.asp"-->

	</body>

</html> 

