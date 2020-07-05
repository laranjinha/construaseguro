<!--#include file="funcoes.asp"-->
<!--#include file="inicial.asp"-->
<%
   Session("gvMenu")="0"
   
   'Registro entrada no sistema
   IF Session("gvAtividadePF")="" OR Session("gvAtividadePF")<>"CADASTRO PF" THEN
      Session("gvAtividadePF")="CADASTRO PF" 
   end if

   'Verifico o protocolo
   IF len(trim(Session("gvProtocolo")))=0 then
      response.redirect("sistemas.asp") 
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
			var ida_nome=document.form.textida_nome.value;
			var ida_email=document.form.textida_email.value;
			var idp_nome=document.form.textidp_nome.value;
			var idp_email=document.form.textidp_email.value;
			var idp_fone=document.form.textidp_fone.value;
			var idp_sexo=document.form.selectidp_sexo.value;
			var idp_idade=document.form.textidp_idade.value;
			var idp_profissao=document.form.textidp_profissao.value;
			var idp_cpf=document.form.textidp_cpf.value;
			var idp_endereco=document.form.textidp_endereco.value;
			var idp_numero=document.form.textidp_numero.value;
			var idp_bairro=document.form.textidp_bairro.value;
			var idp_municipio=document.form.textidp_municipio.value;
			var idp_estado=document.form.textidp_estado.value;
			var idp_cep=document.form.textidp_cep.value;
			//var idp_ref=document.form.textidp_ref.value;
			//var idp_link=document.form.textidp_link.value;
			//var idp_aceite=document.form.selectidp_aceite.value;
			var ido_endereco=document.form.textido_endereco.value;
			var ido_numero=document.form.textido_numero.value;
			var ido_bairro=document.form.textido_bairro.value;
			var ido_municipio=document.form.textido_municipio.value;
			var ido_estado=document.form.textido_estado.value;
			var ido_cep=document.form.textido_cep.value;
			//var ido_ref=document.form.textido_ref.value;
			var ido_desc_obra=document.form.selectido_desc_obra.value;
			var ido_porte_obra=document.form.selectido_porte_obra.value;
			var ido_nr_func=document.form.textido_nr_func.value;
			var ido_prev_dtin=document.form.textido_prev_dtin.value;
			var ido_prev_dtfi=document.form.textido_prev_dtfi.value;
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
				if (ida_nome.length==0)
				{
					alert("Nome não pode ficar em branco!");
					document.form.textida_nome.focus();
					vbRetorno=false;
				}
			}
			
			if (vbRetorno)
			{
				if (ida_email.length==0)
				{
					alert("E-mail não pode ficar em branco!");
					document.form.textida_email.focus();
					vbRetorno=false;
				}
				else
				{
					if (!ValidaMail(ida_email))
					{
						alert("E-mail inválido!");
						document.form.textida_email.focus();
						vbRetorno=false;						
					}
				}
			}

			if (vbRetorno)
			{
				if (idp_nome.length==0)
				{
					alert("Nome não pode ficar em branco!");
					document.form.textidp_nome.focus();
					vbRetorno=false;
				}
			}	

			if (vbRetorno)
			{
				if (idp_email.length==0)
				{
					alert("E-mail não pode ficar em branco!");
					document.form.textidp_email.focus();
					vbRetorno=false;
				}
				else
				{
					if(!ValidaMail(idp_email))
					{
						alert("E-mail inválido!");
						document.form.textidp_email.focus();
						vbRetorno=false;						
					}
				}
			}			

			if (vbRetorno)
			{
				if (idp_fone.length==0)
				{
					alert("Telefone não pode ficar em branco!");
					document.form.textidp_fone.focus();
					vbRetorno=false;
				}
			}			

			if (vbRetorno)
			{
				if (idp_sexo.length==0)
				{
					alert("Sexo não pode ficar em branco!");
					document.form.selectidp_sexo.focus();
					vbRetorno=false;
				}
			}	

			if (vbRetorno)
			{
				if (idp_idade.length==0)
				{
					alert("Idade não pode ficar em branco!");
					document.form.textidp_idade.focus();
					vbRetorno=false;
				}
			}			
			
			if (vbRetorno)
			{
				if (idp_profissao.length==0)
				{
					alert("Profissão não pode ficar em branco!");
					document.form.textidp_profissao.focus();
					vbRetorno=false;
				}
			}	

			if (vbRetorno)
			{
				if (idp_cpf.length==0)
				{
					alert("CPF não pode ficar em branco!");
					document.form.textidp_cpf.focus();
					vbRetorno=false;
				}
				else
				{
					if (!valida_cpf(idp_cpf))
					{
						alert("CPF inválido!");
						document.form.textidp_cpf.focus();
						vbRetorno=false;						
					}
				}
			}						

			if (vbRetorno)
			{
				if (idp_endereco.length==0)
				{
					alert("Endereço não pode ficar em branco!");
					document.form.textidp_endereco.focus();
					vbRetorno=false;
				}
			}	

			if (vbRetorno)
			{
				if (idp_numero.length==0)
				{
					alert("Número não pode ficar em branco!");
					document.form.textidp_numero.focus();
					vbRetorno=false;
				}
			}			

			if (vbRetorno)
			{
				if (idp_bairro.length==0)
				{
					alert("Bairro não pode ficar em branco!");
					document.form.textidp_bairro.focus();
					vbRetorno=false;
				}
			}	

			if (vbRetorno)
			{
				if (idp_municipio.length==0)
				{
					alert("Município não pode ficar em branco!");
					document.form.textidp_municipio.focus();
					vbRetorno=false;
				}
			}	

			if (vbRetorno)
			{
				if (idp_estado.length==0)
				{
					alert("Estado não pode ficar em branco!");
					document.form.textidp_estado.focus();
					vbRetorno=false;
				}
			}		

			if (vbRetorno)
			{
				if (idp_cep.length==0)
				{
					alert("Cep não pode ficar em branco!");
					document.form.textidp_cep.focus();
					vbRetorno=false;
				}
			}					

			if (vbRetorno)
			{
				if (ido_endereco.length==0)
				{
					alert("Endereço não pode ficar em branco!");
					document.form.textido_endereco.focus();
					vbRetorno=false;
				}
			}	

			if (vbRetorno)
			{
				if (ido_numero.length==0)
				{
					alert("Número não pode ficar em branco!");
					document.form.textido_numero.focus();
					vbRetorno=false;
				}
			}			

			if (vbRetorno)
			{
				if (ido_bairro.length==0)
				{
					alert("Bairro não pode ficar em branco!");
					document.form.textido_bairro.focus();
					vbRetorno=false;
				}
			}	

			if (vbRetorno)
			{
				if (ido_municipio.length==0)
				{
					alert("Município não pode ficar em branco!");
					document.form.textido_municipio.focus();
					vbRetorno=false;
				}
			}	

			if (vbRetorno)
			{
				if (ido_estado.length==0)
				{
					alert("Estado não pode ficar em branco!");
					document.form.textido_estado.focus();
					vbRetorno=false;
				}
			}		

			if (vbRetorno)
			{
				if (ido_cep.length==0)
				{
					alert("Cep não pode ficar em branco!");
					document.form.textido_cep.focus();
					vbRetorno=false;
				}
			}		

			if (vbRetorno)
			{
				if (ido_desc_obra.length==0)
				{
					alert("Descrição da obra não pode ficar em branco!");
					document.form.selectido_desc_obra.focus();
					vbRetorno=false;
				}
			}		

			if (vbRetorno)
			{
				if (ido_porte_obra.length==0)
				{
					alert("Porte da obra não pode ficar em branco!");
					document.form.selectido_porte_obra.focus();
					vbRetorno=false;
				}
			}						

			if (vbRetorno)
			{
				if (ido_nr_func.length==0)
				{
					alert("Número de Funcionários não pode ficar em branco!");
					document.form.textido_nr_func.focus();
					vbRetorno=false;
				}
			}

			if (vbRetorno)
			{
				if (ido_prev_dtin.length==0)
				{
					alert("Periodo de início e de término não pode ficar em branco!");
					document.form.textido_prev_dtin.focus();
					vbRetorno=false;
				}
				else
				{
					if(!CheckData(ido_prev_dtin))
					{
						alert("Data inválida!");
						document.form.textido_prev_dtin.focus();
						vbRetorno=false;						
					}
				}
			}		

			if (vbRetorno)
			{
				if (ido_prev_dtfi.length==0)
				{
					alert("Periodo de início e de término não pode ficar em branco!");
					document.form.textido_prev_dtfi.focus();
					vbRetorno=false;
				}
				else
				{
					if(!CheckData(ido_prev_dtfi))
					{
						alert("Data inválida!");
						document.form.textido_prev_dtfi.focus();
						vbRetorno=false;						
					}

				}
			}					

			
			//Verifico se esta tudo ok
			if (vbRetorno)
			{
				document.form.submit();
			}	
			
		}
	</SCRIPT>	

	<body onload="horizontal();botaotopo();document.form.textida_nome.focus();">

        <!--#include file="cabecalho.asp"-->

		<main class="principal">

            <!-- Coloco return=false pois quero que chama a rotina do onsubmit. Se deixar ele automaticamente envia o form -->
            <FORM method="post" action="pfB.asp" id="form" name="form" onSubmit="return false;">
			<input type="hidden" name="textid_protocolo" id="textid_protocolo"  value=<%=Session("gvProtocolo")%>>

			<table id="dados">
				<caption class="caption-titulo">
					<div id="caption">
						<img src="img/pf.png" width="95" height="37">
						<h1>Cadastro Pessoa Física</h1>
						<h2>(Permite o Cadastro PF)</h2>
                	</div>
				</caption>

				<thead>
					<td colspan="1" class="thead-info">Informe todos os dados para o cadastro:</td>
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
								<label class="label-dados">Informe os seus dados</label>
							</p>
							<p>
				  				<label class="porc30w">Nome: </label>
				  				<input name="textida_nome" type="text" id="textida_nome" onKeyDown="TABEnter();" size="80" maxlength="80">
			      				* 
	  		    			</p>		
							<p>
				  				<label class="porc30w">E-mail: </label>
				  				<input name="textida_email" type="text" id="textida_email" onKeyDown="TABEnter();" size="80" maxlength="100">
			      				* 
	  		    			</p>	
							<p>
								<label class="label-dados">Informe os dados do proprietário da obra</label>
							</p>
							<p>
				  				<label class="porc30w">Nome do proprietario: </label>
				  				<input name="textidp_nome" type="text" id="idp_nome" onKeyDown="TABEnter();" size="80" maxlength="80">
			      				* 
	  		    			</p>	
							<p>
				  				<label class="porc30w">E-mail: </label>
				  				<input name="textidp_email" type="text" id="textidp_email" onKeyDown="TABEnter();" size="80" maxlength="100">
			      				* 
	  		    			</p>	
							<p>
				  				<label class="porc30w">Telefone para Contato: </label>
				  				<input name="textidp_fone" type="text" id="textidp_fone" onkeypress="return txtBoxFormat(document.form, 'textidp_fone', '99999999999999999999', event);" onKeyDown="TABEnter();" size="20" maxlength="20">
			      				* 
	  		    			</p>
							<p>
				  				<label class="porc30w">Sexo:</label>
				    			<select name="selectidp_sexo" size="1" id="selectidp_sexo" onKeyDown="TABEnter();">
								<option value="" selected>&lt;selecionar&gt;</option>
								<option value="M">Masculino</option>
								<option value="F">Feminino</option>
                    			</select>
								*
				  			</p>
							<p>
				  				<label class="porc30w">Idade: </label>
				  				<input name="textidp_idade" type="text" id="textidp_idade" onkeypress="return txtBoxFormat(document.form, 'textidp_idade', '999', event);" onKeyDown="TABEnter();" size="3" maxlength="3">
			      				* 
	  		    			</p>
							<p>
				  				<label class="porc30w">Profissão: </label>
				  				<input name="textidp_profissao" type="text" id="textidp_profissao" onKeyDown="TABEnter();" size="50" maxlength="50">
			      				* 
	  		    			</p>	
							<p>
				  				<label class="porc30w">CPF: </label>
				  				<input name="textidp_cpf" type="text" id="textidp_cpf" onkeypress="return txtBoxFormat(document.form, 'textidp_cpf', '99999999999', event);" onKeyDown="TABEnter();" size="11" maxlength="11">
			      				* 
	  		    			</p>
							<p>
				  				<label class="porc30w">Endereço: </label>
				  				<input name="textidp_endereco" type="text" id="textidp_endereco" onKeyDown="TABEnter();" size="80" maxlength="80">
			      				* 
	  		    			</p>
							<p>
				  				<label class="porc30w">Número: </label>
				  				<input name="textidp_numero" type="text" id="textidp_numero" onkeypress="return txtBoxFormat(document.form, 'textidp_numero', '999999', event);" onKeyDown="TABEnter();" size="6" maxlength="6">
			      				* 
	  		    			</p>
							<p>
				  				<label class="porc30w">Bairro: </label>
				  				<input name="textidp_bairro" type="text" id="textidp_bairro" onKeyDown="TABEnter();" size="40" maxlength="40">
			      				* 
	  		    			</p>
							<p>
				  				<label class="porc30w">Município: </label>
				  				<input name="textidp_municipio" type="text" id="textidp_municipio" onKeyDown="TABEnter();" size="50" maxlength="50">
			      				* 
	  		    			</p>
							<p>
				  				<label class="porc30w">Estado: </label>
				  				<input name="textidp_estado" type="text" id="textidp_estado" onKeyDown="TABEnter();" size="2" maxlength="2">
			      				* 
	  		    			</p>
							<p>
				  				<label class="porc30w">CEP: </label>
				  				<input name="textidp_cep" type="text" id="textidp_cep" onkeypress="return txtBoxFormat(document.form, 'textidp_cep', '99999999', event);" onKeyDown="TABEnter();" size="8" maxlength="8">
			      				* 
	  		    			</p>
							<p>
				  				<label class="porc30w">Referência (Opcional): </label>
				  				<input name="textidp_ref" type="text" id="textidp_ref" onKeyDown="TABEnter();" size="50" maxlength="50">
	  		    			</p>
 							<p>
				  				<label class="porc30w">Link do Linkedin, Instagram ou Facebook (opcional):</label>
				  				<textarea name="textidp_link" cols="80" rows="3" id="textidp_link" onKeyDown="TABEnter();"></textarea>
 			    			</p>
							<p>
				  				<label class="porc30w">Aceita receber novidades do Sebrae por e-mail:</label>
				    			<select name="selectidp_aceite" size="1" id="selectidp_aceite" onKeyDown="TABEnter();">
								<option value="" selected>&lt;selecionar&gt;</option>
								<option value="S">Sim</option>
								<option value="N">Não</option>
                    			</select>
				  			</p>
							<p>
								<label class="label-dados">Informe os dados da obra</label>
							</p>
							<p>
				  				<label class="porc30w">Endereço: </label>
				  				<input name="textido_endereco" type="text" id="textido_endereco" onKeyDown="TABEnter();" size="80" maxlength="80">
			      				* 
	  		    			</p>
							<p>
				  				<label class="porc30w">Número: </label>
				  				<input name="textido_numero" type="text" id="textido_numero" onkeypress="return txtBoxFormat(document.form, 'textido_numero', '999999', event);" onKeyDown="TABEnter();" size="6" maxlength="6">
			      				* 
	  		    			</p>
							<p>
				  				<label class="porc30w">Bairro: </label>
				  				<input name="textido_bairro" type="text" id="textido_bairro" onKeyDown="TABEnter();" size="40" maxlength="40">
			      				* 
	  		    			</p>
							<p>
				  				<label class="porc30w">Município: </label>
				  				<input name="textido_municipio" type="text" id="textido_municipio" onKeyDown="TABEnter();" size="50" maxlength="50">
			      				* 
	  		    			</p>
							<p>
				  				<label class="porc30w">Estado: </label>
				  				<input name="textido_estado" type="text" id="textido_estado" onKeyDown="TABEnter();" size="2" maxlength="2">
			      				* 
	  		    			</p>
							<p>
				  				<label class="porc30w">CEP: </label>
				  				<input name="textido_cep" type="text" id="textido_cep" onkeypress="return txtBoxFormat(document.form, 'textido_cep', '99999999', event);" onKeyDown="TABEnter();" size="8" maxlength="8">
			      				* 
	  		    			</p>
							<p>
				  				<label class="porc30w">Referência (Opcional): </label>
				  				<input name="textido_ref" type="text" id="textido_ref" onKeyDown="TABEnter();" size="50" maxlength="50">
	  		    			</p>
							<p>
				  				<label class="porc30w">Como você descreveria a sua obra:</label>
				    			<select name="selectido_desc_obra" size="1" id="selectido_desc_obra" onKeyDown="TABEnter();">
								<option value="" selected>&lt;selecionar&gt;</option>
								<%
									cSQL = "SELECT * FROM tb_ido_desc_obra ORDER BY desc_obra"  	
				        			Set oRs = MyConn.Execute(cSQL)
									do while not oRs.eof
										response.write("<OPTION value=" & chr(34) & oRs("ido_desc_obra") & chr(34) & ">" & oRs("desc_obra") & "</OPTION>")
										oRs.movenext
									loop
									oRs.close
									Set oRs=nothing
								%>
                    			</select>
								*
				  			</p>
							<p>
							    <label class="porc30w">&nbsp;</label>
				  				<strong class="fonte">Fonte para consulta: <a href="https://cnae.ibge.gov.br/?view=secao&tipo=cnae&versaoclasse=7&secao=F">https://cnae.ibge.gov.br...</a></strong>
				  			</p>
							<p>
				  				<label class="porc30w">Qual porte da sua obra:</label>
				    			<select name="selectido_porte_obra" size="1" id="selectido_porte_obra" onKeyDown="TABEnter();">
								<option value="" selected>&lt;selecionar&gt;</option>
								<%
								Session.CodePage=65001
									cSQL = "SELECT * FROM tb_ido_porte_obra ORDER BY porte_obra"  	
				        			Set oRs = MyConn.Execute(cSQL)
									do while not oRs.eof
										response.write("<OPTION value=" & chr(34) & oRs("ido_porte_obra") & chr(34) & ">" & oRs("porte_obra") & "</OPTION>")
										oRs.movenext
									loop
									oRs.close
									Set oRs=nothing
								%>
                    			</select>
								*
				  			</p>
							<p>
							    <label class="porc30w">&nbsp;</label>
				  				<strong class="fonte">Fonte para consulta: <a href="https://www.scielo.br/scielo.php?pid=S1678-86212014000300007&script=sci_arttext#:~:text=As%20obras%20de%20pequeno%20porte,edifica%25C3%25A7%25C3%25B5es%20decinco%20a%20quatorze%20pavimentos" target="_blank">https://www.scielo.br...</a></strong>
				  			</p>
							<p>
				  				<label class="porc30w">Quantos funcionários atuarão na sua obra: </label>
				  				<input name="textido_nr_func" type="text" id="textido_nr_func" onkeypress="return txtBoxFormat(document.form, 'textido_nr_func', '9999', event);" onKeyDown="TABEnter();" size="4" maxlength="4">
			      				* 
	  		    			</p>
							<p>
				  				<label class="porc30w">Previsão de início e de término: </label>
				  				<input name="textido_prev_dtin" type="text" id="textido_prev_dtin" onkeypress="return txtBoxFormat(document.form, 'textido_prev_dtin', '99/99/9999', event);" onKeyDown="TABEnter();" size="10" maxlength="10">
								&nbsp;a&nbsp;
								<input name="textido_prev_dtfi" type="text" id="textido_prev_dtfi" onkeypress="return txtBoxFormat(document.form, 'textido_prev_dtfi', '99/99/9999', event);" onKeyDown="TABEnter();" size="10" maxlength="10">
  			      				* 
	  		    			</p>
						 </td>
					</tr>
				</tbody>

				<tfoot>
        			<tr>
                		<td colspan="1" class="tfoot-info">
					      <input type="button" name="Button" value="Gravar e responder o checklist" onClick="javascript:confirmar();">
						</td>
       				</tr>
        			<tr>
                		<td>
							<p>
					    		<b>*</b> Campos obrigat&oacute;rios <br>
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
