<!--#include file="funcoes.asp"-->
<!--#include file="inicial.asp"-->
<%

    'verifico se esta na opção corrreta
    if Session("gvAtividadePF")="CADASTRO PF" then
      'ok
    else
      Response.Redirect("sistemas.asp")
    end if

    Session("gvIDPF")=""
   
    'pego os campos 
    'mid_pf=trim(Request.form("textid_pf"))
    mid_protocolo=trim(Request.form("textid_protocolo"))
    mida_nome=trim(Request.form("textida_nome"))
    mida_email=trim(Request.form("textida_email"))
    midp_nome=trim(Request.form("textidp_nome"))
    midp_email=trim(Request.form("textidp_email"))
    midp_fone=trim(Request.form("textidp_fone"))
    midp_sexo=trim(Request.form("selectidp_sexo"))
    midp_idade=trim(Request.form("textidp_idade"))
    midp_profissao=trim(Request.form("textidp_profissao"))
    midp_cpf=trim(Request.form("textidp_cpf"))
    midp_endereco=trim(Request.form("textidp_endereco"))
    midp_numero=trim(Request.form("textidp_numero"))
    midp_bairro=trim(Request.form("textidp_bairro"))
    midp_municipio=trim(Request.form("textidp_municipio"))
    midp_estado=trim(Request.form("textidp_estado"))
    midp_cep=trim(Request.form("textidp_cep"))
    midp_ref=trim(Request.form("textidp_ref"))
    midp_link=trim(Request.form("textidp_link"))
    midp_aceite=trim(Request.form("selectidp_aceite"))
    mido_endereco=trim(Request.form("textido_endereco"))
    mido_numero=trim(Request.form("textido_numero"))
    mido_bairro=trim(Request.form("textido_bairro"))
    mido_municipio=trim(Request.form("textido_municipio"))
    mido_estado=trim(Request.form("textido_estado"))
    mido_cep=trim(Request.form("textido_cep"))
    mido_ref=trim(Request.form("textido_ref"))
    mido_desc_obra=trim(Request.form("selectido_desc_obra"))
    mido_porte_obra=trim(Request.form("selectido_porte_obra"))
    mido_nr_func=trim(Request.form("textido_nr_func"))
    mido_prev_dtin=trim(Request.form("textido_prev_dtin"))
    mido_prev_dtfi=trim(Request.form("textido_prev_dtfi"))

    //mid_pf=EvitarCmdSql(mid_pf)
    mid_protocolo=EvitarCmdSql(mid_protocolo)
    mida_nome=EvitarCmdSql(mida_nome)
    mida_email=EvitarCmdSql(mida_email)
    midp_nome=EvitarCmdSql(midp_nome)
    midp_email=EvitarCmdSql(midp_email)
    midp_fone=EvitarCmdSql(midp_fone)
    midp_sexo=EvitarCmdSql(midp_sexo)
    midp_idade=EvitarCmdSql(midp_idade)
    midp_profissao=EvitarCmdSql(midp_profissao)
    midp_cpf=EvitarCmdSql(midp_cpf)
    midp_endereco=EvitarCmdSql(midp_endereco)
    midp_numero=EvitarCmdSql(midp_numero)
    midp_bairro=EvitarCmdSql(midp_bairro)
    midp_municipio=EvitarCmdSql(midp_municipio)
    midp_estado=EvitarCmdSql(midp_estado)
    midp_cep=EvitarCmdSql(midp_cep)
    midp_ref=EvitarCmdSql(midp_ref)
    midp_link=EvitarCmdSql(midp_link)
    midp_aceite=EvitarCmdSql(midp_aceite)
    mido_endereco=EvitarCmdSql(mido_endereco)
    mido_numero=EvitarCmdSql(mido_numero)
    mido_bairro=EvitarCmdSql(mido_bairro)
    mido_municipio=EvitarCmdSql(mido_municipio)
    mido_estado=EvitarCmdSql(mido_estado)
    mido_cep=EvitarCmdSql(mido_cep)
    mido_ref=EvitarCmdSql(mido_ref)
    mido_desc_obra=EvitarCmdSql(mido_desc_obra)
    mido_porte_obra=EvitarCmdSql(mido_porte_obra)
    mido_nr_func=EvitarCmdSql(mido_nr_func)
    mido_prev_dtin=EvitarCmdSql(mido_prev_dtin)
    mido_prev_dtfi=EvitarCmdSql(mido_prev_dtfi)

%>
<!--#include file="cnxSQLServer.asp"-->
<% 
	on error resume next
	err.number=0

    'Verifico se não existe o cadastro
    'cSQL = "SELECT count(*) as total FROM tp_pf WHERE ...."
	'Set oRs = MyConn.Execute(cSQL)
    'if clng(oRs("total"))>0 then
    '   oRs.close
    '   set oRs=nothing
	'   MyConn.close
	'   set MyConn=nothing
	'   Response.Redirect "pfA.asp?ret=Já existe um registro com .... Não posso duplicar registro."
    'end if
    'oRs.close
    'set oRs=nothing

	'incluir o registro
	MyConn.Begintrans

	'vou incluir a caixa
	if err.number=0 then
   	    cSQL = "INSERT INTO tb_pf(id_protocolo, ida_nome, ida_email, idp_nome, idp_email, idp_fone, idp_sexo, idp_idade, "
        cSQL = cSQL & "idp_profissao, idp_cpf, idp_endereco, idp_numero, idp_bairro, idp_municipio, idp_estado, idp_cep, idp_ref, idp_link, "
        cSQL = cSQL & "idp_aceite, ido_endereco, ido_numero, ido_bairro, ido_municipio, ido_estado, ido_cep, ido_ref, ido_desc_obra, ido_porte_obra, "
        cSQL = cSQL & " ido_nr_func, ido_prev_dtin, ido_prev_dtfi) VALUES("
        'cSQL = cSQL & mid_pf & ","
        cSQL = cSQL & mid_protocolo & ","
        cSQL = cSQL & "'" & mida_nome & "',"
        cSQL = cSQL & "'" & mida_email & "',"
        cSQL = cSQL & "'" & midp_nome & "',"
        cSQL = cSQL & "'" & midp_email & "',"
        cSQL = cSQL & "'" & midp_fone & "',"
        cSQL = cSQL & "'" & midp_sexo & "',"
        cSQL = cSQL & midp_idade & ","
        cSQL = cSQL & "'" & midp_profissao & "',"
        cSQL = cSQL & "'" & midp_cpf & "',"
        cSQL = cSQL & "'" & midp_endereco & "',"
        cSQL = cSQL & midp_numero & ","
        cSQL = cSQL & "'" & midp_bairro & "',"
        cSQL = cSQL & "'" & midp_municipio & "',"
        cSQL = cSQL & "'" & midp_estado & "',"
        cSQL = cSQL & "'" & midp_cep & "',"
        cSQL = cSQL & "'" & midp_ref & "',"
        cSQL = cSQL & "'" & mid(midp_link,1,200) & "',"
        cSQL = cSQL & "'" & iif(len(trim(midp_aceite))=0,"N",midp_aceite) & "',"
        cSQL = cSQL & "'" & mido_endereco & "',"
        cSQL = cSQL & mido_numero & ","
        cSQL = cSQL & "'" & mido_bairro & "',"
        cSQL = cSQL & "'" & mido_municipio & "',"
        cSQL = cSQL & "'" & mido_estado & "',"
        cSQL = cSQL & "'" & mido_cep & "',"
        cSQL = cSQL & "'" & mido_ref & "',"
        cSQL = cSQL & mido_desc_obra & ","
        cSQL = cSQL & mido_porte_obra & ","
        cSQL = cSQL & mido_nr_func & ","
        cSQL = cSQL & "'" & DataMySQL(mido_prev_dtin) & "',"
        cSQL = cSQL & "'" & DataMySQL(mido_prev_dtfi) & "'"
   	    cSQL = cSQL & ")"
  
        MyConn.Execute(cSQL)
    end if  

    'Pego o id gerado 	
	if err.number=0 then
	   cSQL = "SELECT LAST_INSERT_ID() AS NewID"
	   Set oRs=MyConn.Execute(cSQL)
	   mid_pf=trim(cstr(oRs("NewID")))
	   oRs.close
   	   Set oRs=nothing		
  	end if

	'Verifico se pegou o id do registro criado
	if err.number=0 then
		if len(trim(mid_pf))=0 or isnull(mid_pf) then
		   err.number="999"
		   err.description="Não conseguiu pegar o ID da Inclusãodo registro!"
        else
           Session("gvIDPF")=mid_pf
		end if     
	end if

	if err.number=0 then
	   MyConn.CommitTrans
		   
       MyConn.close
  	   set MyConn=nothing				
		   
	   'vou para tela do checklist
	   Response.Redirect "checkA.asp?ret=0"
	else
       MyConn.RollbackTrans
    
       MyConn.close
       set MyConn=nothing				

  	   'volto a tela pai
	   Response.Redirect "pfA.asp?ret=Inclusão Invalida!" & abs(err.number) & "-" & replace(err.description,"'"," ")
	end if
%>
