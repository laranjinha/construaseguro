<!--#include file="funcoes.asp"-->
<%
    'Expirar pagina ao enviar.
    response.buffer=true
    response.expires = 0
    response.expiresabsolute = Now() - 3
    response.addHeader "pragma","no-cache"
    response.addHeader "cache-control","private"
    Response.CacheControl = "no-cache"

    'Pego os dados
    numRET=trim(request.QueryString("ret"))
	
    'evitar sql injection
    numRET=EvitarCmdSql(numRET)	
   
    'Verifico se tem algum problema
    mAviso=""
    if len(trim(numRET))>0 and trim(numRET)<>"0" then
       mAviso=trim(numRET)
    end if	

    'Zero variaveis de sessão
   	Session("gvMenu")="0"

    Session.LCID=1046				' Data no formato brasileiro	
	Session("LoginAtivo") = 0       ' Indica se foi realizado o login do usuario 
	Session("gvMenu")=""            ' Opção do menu selecionada   
	Session("gvSubMenu")=""         ' Opção do menu selecionada    
	Session("gvPagina")=0			' numero da pagina atual da grade
    Session("gvUsuario")=""         ' login 
    Session("gvUsuarioID")=0        ' id do usuario   
    Session("gvNomeUsuario")=""	    ' nome do usuario
	Session("gvPerfil")=""	        ' perfil
	Session("gvHUA")=""				' controle de roubo
	Session("gvLA")=""
	Session("gvRA")=""
	Session("gvHC")=""
%>
<!DOCTYPE html>
<html lang="pt-br">

	<head>
        <!--#include file="head.asp"-->
	</head>

	<script language="JavaScript">
		function VoltarOpc()
		{
			js_a='sair.asp';
    		location.href=js_a;	
		}

		function confirmadados()
		{
			var vbRetorno=true;
			
			js_log = document.form1.txtlog.value;
			js_sen = document.form1.txtsen.value;
	
		    if(vbRetorno)
			{
				if(js_log.length == 0)
				{
					alert("Informe seu login!")
					document.form1.txtlog.focus();
					vbRetorno=false;
				}	
			}

		    if(vbRetorno)
			{
				if(js_sen.length == 0)
				{
					alert("Informe sua senha!")
					document.form1.txtsen.focus();
					vbRetorno=false;
				}	
			}
			
			//Verifico se esta tudo ok
			if (vbRetorno)
			{
				document.form1.submit();
			}
			else
			{
				return vbRetorno;	
			}	
		}
	</script>

	<body  onload="document.form1.txtlog.focus();">

        <!--#include file="cabecalho.asp"-->

		<main class="principal">

            <!-- Coloco return=false pois quero que chama a rotina do onsubmit. Se deixar ele automaticamente envia o form -->
            <form method="post" action="logB.asp" id="form1" name="form1" onSubmit="return false;">

			<table id="dados">
				<caption class="caption-titulo">
					<div id="caption">
						<img src="img/login.png" width="95" height="37">
						<h1>Login </h1>
						<h2>(Informe seu login e senha para acessar o portal)</h2>
                	</div>
				</caption>

				<thead>
					<td colspan="1" class="thead-info">Informe seu login:</td>
				</thead>

				<tbody>
					<tr>
						<td>
							<%
								'Verifico se teve algum erro 
								if len(trim(mAviso))>0 then
									response.write("<p class=""label-titulo-msg"">")
   									response.write(mAviso)
   									response.write("</p>")
								end if
							%> 
                			<p>
                  				<label>Login:</label>
                  				<input name="txtlog" type="text" id="txtlog" size="15" maxlength="15" onKeyDown="TABEnter();"> 
                  				* 
                			</p>
                			<p>
                  				<label>Senha:</label>
                  				<input name="txtsen" type="password" id="txtsen" size="15" maxlength="15" onKeyDown="TABEnter();" >
								*                			
							</p>
						 </td>
					</tr>
				</tbody>

				<tfoot>
        			<tr>
                		<td colspan="1" class="tfoot-info">
					      <input type="button" name="Button" value="OK" onClick="javascript:confirmadados();">
						  &nbsp;&nbsp;
						  <input type="button" name="Button2" value="CANCELAR" onClick="javascript:VoltarOpc();">
						</td>
       				</tr>
					<tr>
						<td>
							<p>
								<b>*</b> Campos obrigatórios <br>
							</p>
						</td>
					</tr>
				</tfoot>

			</table>

	        </form>

        </main>

		<!--#include file="rodape.asp"-->

	</body>

</html>
<%
	'mensagem de erro
	if len(trim(mAviso))>0 then
		Response.Write("<script>")
		Response.Write("alert('"&mAviso&"');")
		Response.Write("</script>")	  
	end if  
%> 
