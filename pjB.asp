<!--#include file="funcoes.asp"-->
<!--#include file="inicial.asp"-->
<%

    'verifico se esta na opção corrreta
    if Session("gvAtividadePJ")="CADASTRO PJ" then
      'ok
    else
      Response.Redirect("sistemas.asp")
    end if

    Session("gvIDPJ")=""
   
    'pego os campos 
	//mid_pj=trim(Request.form("textid_pj"))
	mid_protocolo=trim(Request.form("textid_protocolo"))
	mida_nome=trim(Request.form("textida_nome"))
	mida_email=trim(Request.form("textida_email"))
	mide_razao=trim(Request.form("textide_razao"))
	mide_cnpj=trim(Request.form("textide_cnpj"))
	mide_porte=trim(Request.form("selectide_porte"))
	mide_resp=trim(Request.form("textide_resp"))
	mide_email_contato=trim(Request.form("textide_email_contato"))
	mide_fone_contato=trim(Request.form("textide_fone_contato"))
	mide_endereco=trim(Request.form("textide_endereco"))
	mide_numero=trim(Request.form("textide_numero"))
	mide_bairro=trim(Request.form("textide_bairro"))
    mide_municipio=trim(Request.form("textide_municipio"))
	mide_estado=trim(Request.form("textide_estado"))
	mide_cep=trim(Request.form("textide_cep"))
	//mide_ref=trim(Request.form("textide_ref"))
	//mide_link=trim(Request.form("textide_link"))
	//mide_aceite=trim(Request.form("selectide_aceite"))
	mido_endereco=trim(Request.form("textido_endereco"))
	mido_numero=trim(Request.form("textido_numero"))
	mido_bairro=trim(Request.form("textido_bairro"))
	mido_municipio=trim(Request.form("textido_municipio"))
	mido_estado=trim(Request.form("textido_estado"))
	mido_cep=trim(Request.form("textido_cep"))
	//mido_ref=trim(Request.form("textido_ref"))
	mido_desc_obra=trim(Request.form("selectido_desc_obra"))
	mido_porte_obra=trim(Request.form("selectido_porte_obra"))
	mido_nr_func=trim(Request.form("textido_nr_func"))
	mido_prev_dtin=trim(Request.form("textido_prev_dtin"))
	mido_prev_dtfi=trim(Request.form("textido_prev_dtfi"))

	//mid_pj=EvitarCmdSql(mid_pj)
	mid_protocolo=EvitarCmdSql(mid_protocolo)
	mida_nome=EvitarCmdSql(mida_nome)
	mida_email=EvitarCmdSql(mida_email)
	mide_razao=EvitarCmdSql(mide_razao)
	mide_cnpj=EvitarCmdSql(mide_cnpj)
	mide_porte=EvitarCmdSql(mide_porte)
	mide_resp=EvitarCmdSql(mide_resp)
	mide_email_contato=EvitarCmdSql(mide_email_contato)
	mide_fone_contato=EvitarCmdSql(mide_fone_contato)
	mide_endereco=EvitarCmdSql(mide_endereco)
	mide_numero=EvitarCmdSql(mide_numero)
	mide_bairro=EvitarCmdSql(mide_bairro)
    mide_municipio=EvitarCmdSql(mide_municipio)
	mide_estado=EvitarCmdSql(mide_estado)
	mide_cep=EvitarCmdSql(mide_cep)
	//mide_ref=EvitarCmdSql(mide_ref)
	//mide_link=EvitarCmdSql(mide_link)
	//mide_aceite=EvitarCmdSql(mide_aceite)
	mido_endereco=EvitarCmdSql(mido_endereco)
	mido_numero=EvitarCmdSql(mido_numero)
	mido_bairro=EvitarCmdSql(mido_bairro)
	mido_municipio=EvitarCmdSql(mido_municipio)
	mido_estado=EvitarCmdSql(mido_estado)
	mido_cep=EvitarCmdSql(mido_cep)
	//mido_ref=EvitarCmdSql(mido_ref)
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
    'cSQL = "SELECT count(*) as total FROM tp_pj WHERE ...."
	'Set oRs = MyConn.Execute(cSQL)
    'if clng(oRs("total"))>0 then
    '   oRs.close
    '   set oRs=nothing
	'   MyConn.close
	'   set MyConn=nothing
	'   Response.Redirect "pjA.asp?ret=Já existe um registro com .... Não posso duplicar registro."
    'end if
    'oRs.close
    'set oRs=nothing

	'incluir o registro
	MyConn.Begintrans

	'vou incluir a caixa
	if err.number=0 then
   	    cSQL = "INSERT INTO tb_pj(id_protocolo, ida_nome, ida_email, ide_razao, ide_cnpj, ide_porte, ide_resp, ide_email_contato, "
        cSQL = cSQL & "ide_fone_contato, ide_endereco, ide_numero, ide_bairro, ide_municipio, ide_estado, ide_cep, ide_ref, ide_link, ide_aceite, ido_endereco, "
        cSQL = cSQL & "ido_numero, ido_bairro, ido_municipio, ido_estado, ido_cep, ido_ref, ido_desc_obra, ido_porte_obra, ido_nr_func, ido_prev_dtin, ido_prev_dtfi) VALUES("
        'cSQL = cSQL & mid_pj & ","
        cSQL = cSQL & mid_protocolo & ","
        cSQL = cSQL & "'" & mida_nome & "',"
        cSQL = cSQL & "'" & mida_email & "',"
        cSQL = cSQL & "'" & mide_razao & "',"
        cSQL = cSQL & "'" & mide_cnpj & "',"
        cSQL = cSQL & "'" & mide_porte & "',"
        cSQL = cSQL & "'" & mide_resp & "',"
        cSQL = cSQL & "'" & mide_email_contato & "',"
        cSQL = cSQL & "'" & mide_fone_contato & "',"
        cSQL = cSQL & "'" & mide_endereco & "',"
        cSQL = cSQL & mide_numero & ","
        cSQL = cSQL & "'" & mide_bairro & "',"
        cSQL = cSQL & "'" & mide_municipio & "',"
        cSQL = cSQL & "'" & mide_estado & "',"
        cSQL = cSQL & "'" & mide_cep & "',"
        cSQL = cSQL & "'" & mide_ref & "',"
        cSQL = cSQL & "'" & mid(mide_link,1,200) & "',"
        cSQL = cSQL & "'" & iif(len(trim(mide_aceite))=0,"N",mide_aceite) & "',"
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
	   mid_pj=trim(cstr(oRs("NewID")))
	   oRs.close
   	   Set oRs=nothing		
  	end if

	'Verifico se pegou o id do registro criado
	if err.number=0 then
		if len(trim(mid_pj))=0 or isnull(mid_pj) then
		   err.number="999"
		   err.description="Não conseguiu pegar o ID da Inclusãodo registro!"
        else
           Session("gvIDPJ")=mid_pj
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
	   Response.Redirect "pjA.asp?ret=Inclusão Invalida!" & abs(err.number) & "-" & replace(err.description,"'"," ")
	end if
%>
