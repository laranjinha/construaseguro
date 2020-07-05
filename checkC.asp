<!--#include file="funcoes.asp"-->
<!--#include file="inicial.asp"-->
<%
    Session("gvMenu")="0"
   
    'verifico se esta na opção corrreta
	if Session("gvIDPJPF")="PF" then
 	   if Session("gvAtividadePF")="CADASTRO PF" then
 	     'ok
 	   else
	      Response.Redirect("sistemas.asp")
	   end if
	else
 	   if Session("gvAtividadePJ")="CADASTRO PJ" then
 	     'ok
 	   else
	      Response.Redirect("sistemas.asp")
	   end if	
	end if

    'Verifico o protocolo
    IF len(trim(Session("gvProtocolo")))=0 then
      response.redirect("sistemas.asp") 
    end if
%>
<!--#include file="cnxSQLServer.asp"-->
<!DOCTYPE html>
<html lang="pt-br">

	<head>
        <!--#include file="head.asp"-->
	</head>

	<SCRIPT language="JavaScript">	
        function VoltarClick()
        {
            a="portal.asp";
            location.href=a
        }
	</SCRIPT>	

	<body onload="horizontal();botaotopo();">

        <!--#include file="cabecalho.asp"-->

		<main class="principal">

            <!-- Coloco return=false pois quero que chama a rotina do onsubmit. Se deixar ele automaticamente envia o form -->
            <FORM method="post" id="form" name="form" onSubmit="return false;">

			<table id="dados">
			    <%if Session("gvIDPJPF")="PF" then%>
					<caption class="caption-titulo">
						<div id="caption">
							<img src="img/pf.png" width="95" height="37">
							<h1>Cadastro Pessoa Física</h1>
							<h2>(Permite o Cadastro PF)</h2>
						</div>
					</caption>
				<%else%>
					<caption class="caption-titulo">
						<div id="caption">
							<img src="img/pj.png" width="95" height="37">
							<h1>Cadastro Pessoa Jurídica</h1>
							<h2>(Permite o Cadastro PJ)</h2>
						</div>
					</caption>				
				<%end if%>

				<thead>
					<td colspan="1" class="thead-info">
                        <img src="img/obrigado.png">
                    </td>
				</thead>

				<tbody>
					<tr>
						<td>
    						<p>
								<label class="label_pergmsg">
                                    Protocolo Número: <strong style="color:black;"><%=Session("gvProtocolo")%></strong></br>
		                            Total de pontos até o momento: <strong style="color:black;"><%=Session("gvSoma")%></strong> pontos</br>
		                            Esses pontos podem ser convertidos em desconto nas lojas parceiras.</br>
								</label>
							</p>
						 </td>
					</tr>
				</tbody>

				<tfoot>
        			<tr>
                		<td colspan="1" class="tfoot-info">
					      <input type="button" name="Button" value="Fechar" onClick="javascript:VoltarClick();">
						</td>
       				</tr>
				</tfoot>

			</table>	

	        </form>

        </main>

		<!--#include file="rodape.asp"-->

	</body>

</html>
