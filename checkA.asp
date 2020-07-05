<!--#include file="funcoes.asp"-->
<!--#include file="inicial.asp"-->
<%
    Session("gvMenu")="0"
	Session("gvSoma")="0"
   
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

   'Verifico o id do cadastro
   if Session("gvIDPJPF")="PF" then   
      IF len(trim(Session("gvIDPF")))=0 then
         response.redirect("sistemas.asp") 
      end if   
   else
      IF len(trim(Session("gvIDPJ")))=0 then
         response.redirect("sistemas.asp") 
      end if    
   end if

   'Verifico se houve algum retorno
   mRetorno = Trim(Request.QueryString("ret"))
   if len(trim(mRetorno))>0 and mRetorno<>"0" then
      mAviso=EvitarCmdSql(mRetorno)
   else
      mAviso=""	  
   end if 

   'pego a data do sistema
   mData=PegaDataSystema()

%>
<!--#include file="cnxSQLServer.asp"-->
<!DOCTYPE html>
<html lang="pt-br">

	<head>
        <!--#include file="head.asp"-->
	</head>

	<SCRIPT language="JavaScript">	
		function confirmar()
		{
			//var id_pf=document.form.textid_pf.value;
			var id_protocolo=document.form.textid_protocolo.value;
			var fone=document.form.textfone.value;
			var sf2=document.form.selectsf2.value;
			var sf3=document.form.selectsf3.value;
			var sf41=document.form.opc1.checked;
			var sf42=document.form.opc2.checked;
			var sf43=document.form.opc3.checked;
			var sf5=document.form.selectsf5.value;

			var so1=document.form.selectso1.value;
			var so2=document.form.selectso2.value;
			var so3=document.form.selectso3.value;
			var so4=document.form.selectso4.value;
			var so5=document.form.selectso5.value;

			var sp1=document.form.selectsp1.value;
			var sp2=document.form.selectsp2.value;
			var sp3=document.form.selectsp3.value;
			var sp4=document.form.selectsp4.value;
			var sp5=document.form.selectsp5.value;
			var vbRetorno=true;

			if (vbRetorno)
			{
				if (id_protocolo.length==0)
				{
					alert("Protocolo não pode ficar em branco!");
					//document.form.textid_protocolo.focus();
					vbRetorno=false;
				}
			}			
	
			if (vbRetorno)
			{
				if (fone.length==0)
				{
					alert("Pergunta 1A não pode ficar em branco!");
					document.form.textfone.focus();
					vbRetorno=false;
				}
				else
				{
					if (!checkFones(fone))
					{
						alert("Telefone inválido! Lembre-se digitar DDD+num.celular tudo junto e só colocar um novo espaço para separ um novo!");
						document.form.textfone.focus();
					    vbRetorno=false;					
					}
				}
			}
			
			if (vbRetorno)
			{
				if (sf2.length==0)
				{
					alert("Pergunta 2A não pode ficar em branco!");
					document.form.selectsf2.focus();
					vbRetorno=false;
				}
			}

			if (vbRetorno)
			{
				if (sf3.length==0)
				{
					alert("Pergunta 3A não pode ficar em branco!");
					document.form.selectsf3.focus();
					vbRetorno=false;
				}
			}		

			if (vbRetorno)
			{
				if (sf41==false && sf42==false && sf43==false)
				{
					alert("Pergunta 4A não pode ficar em branco!");
					document.form.opc1.focus();
					vbRetorno=false;
				}
			}		

			if (vbRetorno)
			{
				if (sf5.length==0)
				{
					alert("Pergunta 5A não pode ficar em branco!");
					document.form.selectsf5.focus();
					vbRetorno=false;
				}
			}			

			if (vbRetorno)
			{
				if (so1.length==0)
				{
					alert("Pergunta 1B não pode ficar em branco!");
					document.form.selectso1.focus();
					vbRetorno=false;
				}
			}			

			if (vbRetorno)
			{
				if (so2.length==0)
				{
					alert("Pergunta 2B não pode ficar em branco!");
					document.form.selectso2.focus();
					vbRetorno=false;
				}
			}	

			if (vbRetorno)
			{
				if (so3.length==0)
				{
					alert("Pergunta 3B não pode ficar em branco!");
					document.form.selectso3.focus();
					vbRetorno=false;
				}
			}	

			if (vbRetorno)
			{
				if (so4.length==0)
				{
					alert("Pergunta 4B não pode ficar em branco!");
					document.form.selectso4.focus();
					vbRetorno=false;
				}
			}			

			if (vbRetorno)
			{
				if (so5.length==0)
				{
					alert("Pergunta 5B não pode ficar em branco!");
					document.form.selectso5.focus();
					vbRetorno=false;
				}
			}																			

			if (vbRetorno)
			{
				if (sp1.length==0)
				{
					alert("Pergunta 1C não pode ficar em branco!");
					document.form.selectsp1.focus();
					vbRetorno=false;
				}
			}	

			if (vbRetorno)
			{
				if (sp2.length==0)
				{
					alert("Pergunta 2C não pode ficar em branco!");
					document.form.selectsp2.focus();
					vbRetorno=false;
				}
			}			

			if (vbRetorno)
			{
				if (sp3.length==0)
				{
					alert("Pergunta 3C não pode ficar em branco!");
					document.form.selectsp3.focus();
					vbRetorno=false;
				}
			}		

			if (vbRetorno)
			{
				if (sp4.length==0)
				{
					alert("Pergunta 4C não pode ficar em branco!");
					document.form.selectsp4.focus();
					vbRetorno=false;
				}
			}		

			if (vbRetorno)
			{
				if (sp5.length==0)
				{
					alert("Pergunta 5C não pode ficar em branco!");
					document.form.selectsp5.focus();
					vbRetorno=false;
				}
			}											
		
			//Verifico se esta tudo ok
			if (vbRetorno)
			{
				document.form.submit();
			}	
			
		}
	</SCRIPT>	

	<body onload="horizontal();botaotopo();document.form.textfone.focus();">

        <!--#include file="cabecalho.asp"-->

		<main class="principal">

            <!-- Coloco return=false pois quero que chama a rotina do onsubmit. Se deixar ele automaticamente envia o form -->
            <FORM method="post" action="checkB.asp" id="form" name="form" onSubmit="return false;">
			<input type="hidden" name="textid_protocolo" id="textid_protocolo"  value=<%=Session("gvProtocolo")%>>

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
					<td colspan="1" class="thead-info">Informe todos as perguntas para finalizar seu cadastro:</td>
				</thead>

				<tbody>
					<tr>
						<td>
							<%
								'Verifico se teve algum erro 
								if len(trim(maviso))>0 then
									response.write("<p class=""label-titulo-msg"">")
   									response.write(maviso)
   									response.write("</p>")
								end if
							%>  	
							<p>
								<label class="label_pergmsg">
										Este é um checklist que o Sebrae preparou para que você e seus
										colaboradores possam retomar as suas atividades com mais
										transparência, tranquilidade e segurança.
										Agora é a hora de você mostrar para todo mundo que está colocando todo o conhecimento que
										você adquiriu em prática no seu empreendimento. Mãos a obra, então.
								</label>
							</p>
							<p>
								<label class="label-dadosp">Sobre os funcionários</label>
							</p>							
							<p>
				  				<label class="label_perg">1A. Cadastre o DDD e número do celular dos seus funcionários para que o Parceiro de Obra (um robô do Telegram) possa monitorar em tempo real a situação dos seus colaboradores. <strong style="color:red;"></br>ATENÇÃO: Preencha o DDD + número do celular sem espaços, só coloque espaço para separar um novo celular: Ex.191111111111 1922222222</strong> </label>
								&nbsp;<textarea name="textfone" cols="100" rows="4" id="textfone" onKeyDown="TABEnter();"></textarea>  
	  		    			</p>		
							<p>
				  				<label class="label_perg">2A. Seus funcionários receberam orientações sobre as medidas de prevenção passadas pelo Sebrae?</label>
				    			&nbsp;<select name="selectsf2" size="1" id="selectsf2" onKeyDown="TABEnter();">
								<option value="" selected>&lt;selecionar&gt;</option>
								<option value="S">Sim</option>
								<option value="N">Não</option>
                    			</select>
	  		    			</p>								  
							<p>
				  				<label class="label_perg">3A. Seus funcionários foram orientados para não compartilharem utensílios pessoais e a só usarem uniforme no próprio local do trabalho?</label>
				    			&nbsp;<select name="selectsf3" size="1" id="selectsf3" onKeyDown="TABEnter();">
								<option value="" selected>&lt;selecionar&gt;</option>
								<option value="S">Sim</option>
								<option value="N">Não</option>
                    			</select>
	  		    			</p>	
							<p>
				  				<label class="label_perg">4A. Como você tem passado as orientações que aprendeu no Sebrae para os seus funcionários? Marque as opções que mais aplicarem na sua realidade.</label>
								&nbsp;<input type="checkbox" name="opc1" value="1"> Fiz/faço reuniões com meus funcionários com o objetivo de passar todas orientações necessárias</br>
								&nbsp;<input type="checkbox" name="opc2" value="2"> Baixei os materiais de ambientação do Sebrae e estou colocando eles em pontos estratégicos da minha obra</br>
								&nbsp;<input type="checkbox" name="opc3" value="3"> Eu (meu negócio) tenho mantido contato periódico com todos os funcionários usando recursos tecnológicos (Whatsapp, Telegram e etc)</br>
	  		    			</p>	
							<p>
				  				<label class="label_perg">5A. Você tem disponibilizado máscara facial e álcool em gel em quantidade adequada para seus funcionários?</label>
				    			&nbsp;<select name="selectsf5" size="1" id="selectsf5" onKeyDown="TABEnter();">
								<option value="" selected>&lt;selecionar&gt;</option>
								<option value="S">Sim</option>
								<option value="N">Não</option>
                    			</select>
	  		    			</p>								  							  							  
							<p>
								<label class="label-dadosp">Sobre a obra</label>
							</p>	
							<p>
				  				<label class="label_perg">1B. O ambiente de trabalho da obra é a céu aberto ou, caso não seja, está devidamente ventilado?</label>
				    			&nbsp;<select name="selectso1" size="1" id="selectso1" onKeyDown="TABEnter();">
								<option value="" selected>&lt;selecionar&gt;</option>
								<option value="S">Sim</option>
								<option value="N">Não</option>
                    			</select>
	  		    			</p>		
							<p>
				  				<label class="label_perg">2B. As máquinas e ferramentas têm sido higienizadas regularmente?</label>
				    			&nbsp;<select name="selectso2" size="1" id="selectso2" onKeyDown="TABEnter();">
								<option value="" selected>&lt;selecionar&gt;</option>
								<option value="S">Sim</option>
								<option value="N">Não</option>
                    			</select>
	  		    			</p>	
							<p>
				  				<label class="label_perg">3B. Existe algum controle ou restrição para pessoas que não trabalham na obra (fornecedores de materiais e clientes, por exemplo)?</label>
				    			&nbsp;<select name="selectso3" size="1" id="selectso3" onKeyDown="TABEnter();">
								<option value="" selected>&lt;selecionar&gt;</option>
								<option value="S">Sim</option>
								<option value="N">Não</option>
                    			</select>
	  		    			</p>		
							<p>
				  				<label class="label_perg">4B. O canteiro de obra tem sido desinfetado/limpo após a jornada de trabalho dos seus funcionários?</label>
				    			&nbsp;<select name="selectso4" size="1" id="selectso4" onKeyDown="TABEnter();">
								<option value="" selected>&lt;selecionar&gt;</option>
								<option value="S">Sim</option>
								<option value="N">Não</option>
                    			</select>
	  		    			</p>		
							<p>
				  				<label class="label_perg">5B. O ambiente de trabalho respeita o distanciamento mínimo de 1,5m entre cada trabalhador?</label>
				    			&nbsp;<select name="selectso5" size="1" id="selectso5" onKeyDown="TABEnter();">
								<option value="" selected>&lt;selecionar&gt;</option>
								<option value="S">Sim</option>
								<option value="N">Não</option>
                    			</select>
	  		    			</p>	
							<p>
								<label class="label-dadosp">Sobre seu plano de resposta</label>
							</p>	
							<p>
				  				<label class="label_perg">1C. Você consegue saber se um dos seus funcionários está enquadrado no grupo de risco?</label>
				    			&nbsp;<select name="selectsp1" size="1" id="selectsp1" onKeyDown="TABEnter();">
								<option value="" selected>&lt;selecionar&gt;</option>
								<option value="S">Sim</option>
								<option value="N">Não</option>
                    			</select>
	  		    			</p>	
							<p>
				  				<label class="label_perg">2C. Você estabeleceu medidas de identificação e encaminhamento de trabalhadores suspeitos?</label>
				    			&nbsp;<select name="selectsp2" size="1" id="selectsp2" onKeyDown="TABEnter();">
								<option value="" selected>&lt;selecionar&gt;</option>
								<option value="S">Sim</option>
								<option value="N">Não</option>
                    			</select>
	  		    			</p>	
				  				<label class="label_perg">3C. Você preparou alguma espécie de pré-avaliação de saúde na entrada do local de trabalho?</label>
				    			&nbsp;<select name="selectsp3" size="1" id="selectsp3" onKeyDown="TABEnter();">
								<option value="" selected>&lt;selecionar&gt;</option>
								<option value="S">Sim</option>
								<option value="N">Não</option>
                    			</select>
	  		    			</p>	
	  		    			</p>	
				  				<label class="label_perg">4C. Você fez algum escalonamento dos horários das refeições ou, pelo menos, orientou seus trabalhadores a manterem distanciamento social nesse momento?</label>
				    			&nbsp;<select name="selectsp4" size="1" id="selectsp4" onKeyDown="TABEnter();">
								<option value="" selected>&lt;selecionar&gt;</option>
								<option value="S">Sim</option>
								<option value="N">Não</option>
                    			</select>
	  		    			</p>	
	  		    			</p>	
				  				<label class="label_perg">5C. Você consegue monitorar se todas as diretrizes do seu plano de resposta estão conseguindo ser implementadas na sua obra?</label>
				    			&nbsp;<select name="selectsp5" size="1" id="selectsp5" onKeyDown="TABEnter();">
								<option value="" selected>&lt;selecionar&gt;</option>
								<option value="S">Sim</option>
								<option value="N">Não</option>
                    			</select>
	  		    			</p>								  							  						  							  													  							  						  					  							  						
						 </td>
					</tr>
				</tbody>

				<tfoot>
        			<tr>
                		<td colspan="1" class="tfoot-info">
					      <input type="button" name="Button" value="Gravar e finalizar o cadastro" onClick="javascript:confirmar();">
						</td>
       				</tr>
        			<tr>
                		<td>
							<p>
					    		<b> </b> Todos os Campos são obrigat&oacute;rios <br>
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
	MyConn.close
	Set MyConn=nothing

	'mensagem de erro
	if len(trim(mAviso))>0 then
		Response.Write("<script>")
		Response.Write("alert('"&mAviso&"');")
		Response.Write("</script>")	  
	end if  
%> 
