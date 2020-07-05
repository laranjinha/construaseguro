<!--#include file="funcoes.asp"-->
<!--#include file="inicial.asp"-->
<%
   	Session("gvMenu")="0"
	mAchou=0	
%>
<!DOCTYPE html>
<html lang="pt-br">

	<head>
        <!--#include file="head.asp"-->
	</head>

	<script language="JavaScript">
		function VoltarOpc()
		{
			js_a='portal.asp';
    		location.href=js_a;	
		}
	</script>

	<body  onload="horizontal();">

        <!--#include file="cabecalho.asp"-->

		<main class="principal">

            <!-- Coloco return=false pois quero que chama a rotina do onsubmit. Se deixar ele automaticamente envia o form -->
            <form method="post" id="form1" name="form1" onSubmit="return false;">

			<table id="dados">
				<caption class="caption-titulo">
					<div id="caption">
						<img src="img/sistemas.png" width="95" height="37">
						<h1>Cadastros</h1>
						<h2>(Mostra os cadastros em que o usu&aacute;rio tem acesso no momento)</h2>
                	</div>
				</caption>

				<thead>
					  <td colspan="1" class="thead-info">Cadastro para sua escolha:</td>
				</thead>

				<tbody>
					<tr>
						<td align="center">
   							<div class="divsistema">
   								<table>
									<%
										if instr(Session("gvSistemas"),"PF,")>0 then 
									  		mAchou=mAchou+1
     										response.write("<tr>")
       										response.write("<td><a href=""opc.asp?p=2e1"" title=""Clique para acessar o sistema...""><img src=""img/sistemapf.jpg"" onmouseover=this.src=""img/sistemapfup.jpg"" onmouseout=this.src=""img/sistemapf.jpg""></a></td>")
     										response.write("</tr>")
										end if
										
										if instr(Session("gvSistemas"),"PJ,")>0 then 
									  		mAchou=mAchou+1
     										response.write("<tr>")
       										response.write("<td><a href=""opc.asp?p=2x1"" title=""Clique para acessar o sistema...""><img src=""img/sistemapj.jpg"" onmouseover=this.src=""img/sistemapjup.jpg"" onmouseout=this.src=""img/sistemapj.jpg""></a></td>")
     										response.write("</tr>")
										end if

                                        'verifico se tem algum sistema disponivel     
                                    	if mAchou=0 then
											response.write("<tr>")
                         					response.write("<td><strong class=""mensagem"">Você não esta liberado para acessar nenhum sistema neste portal. Fale com o administrador do sistema em caso de duvidas.</strong></td>")
										    response.write("</tr>")
                                    	end if
									%>                         
								</table>
							</div>
						</td>
					</tr>
				</tbody>

				<tfoot>
        			<tr>
                		<td colspan="1" class="tfoot-info">
						  <input type="button" name="Button2" value="VOLTAR" onClick="javascript:VoltarOpc();">
						</td>
       				</tr>
				</tfoot>

			</table>

	        </form>

        </main>

		<!--#include file="rodape.asp"-->

	</body>

</html>
