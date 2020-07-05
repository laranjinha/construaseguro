<!--#include file="funcoes.asp"-->
<!--#include file="inicial.asp"-->
<%

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

    Session("gvSoma")="0"

    'pego os campos 
    mid_protocolo=trim(Request.form("textid_protocolo"))
    mfone=trim(Request.form("textfone"))
    msf2=trim(Request.form("selectsf2"))
    msf3=trim(Request.form("selectsf3"))
    msf41=trim(Request.form("opc1"))
    msf42=trim(Request.form("opc2"))
    msf43=trim(Request.form("opc3"))
    msf5=trim(Request.form("selectsf5"))
    mso1=trim(Request.form("selectso1"))
    mso2=trim(Request.form("selectso2"))
    mso3=trim(Request.form("selectso3"))
    mso4=trim(Request.form("selectso4"))
    mso5=trim(Request.form("selectso5"))
    msp1=trim(Request.form("selectsp1"))
    msp2=trim(Request.form("selectsp2"))
    msp3=trim(Request.form("selectsp3"))
    msp4=trim(Request.form("selectsp4"))
    msp5=trim(Request.form("selectsp5"))    

    mid_protocolo=EvitarCmdSql(mid_protocolo)
    mfone=EvitarCmdSql(mfone)
    msf2=EvitarCmdSql(msf2)
    msf3=EvitarCmdSql(msf3)
    msf41=EvitarCmdSql(msf41)
    msf42=EvitarCmdSql(msf42)
    msf43=EvitarCmdSql(msf43)
    msf5=EvitarCmdSql(msf5)
    mso1=EvitarCmdSql(mso1)
    mso2=EvitarCmdSql(mso2)
    mso3=EvitarCmdSql(mso3)
    mso4=EvitarCmdSql(mso4)
    mso5=EvitarCmdSql(mso5)
    msp1=EvitarCmdSql(msp1)
    msp2=EvitarCmdSql(msp2)
    msp3=EvitarCmdSql(msp3)
    msp4=EvitarCmdSql(msp4)
    msp5=EvitarCmdSql(msp5) 

    'trato opção sf4
    msf4= msf41 & "|" & msf42 & "|" & msf43

    'calculo a soma
    msoma=0
    if (len(mfone)>0) then msoma=msoma+10
    if (msf2="S") then msoma=msoma+10
    if (msf3="S") then msoma=msoma+10
    if (len(trim(msf4))>0) then msoma=msoma+10
    if (msf5="S") then msoma=msoma+10
    if (mso1="S") then msoma=msoma+10
    if (mso2="S") then msoma=msoma+10
    if (mso3="S") then msoma=msoma+10
    if (mso4="S") then msoma=msoma+10
    if (mso5="S") then msoma=msoma+10
    if (msp1="S") then msoma=msoma+10
    if (msp2="S") then msoma=msoma+10
    if (msp3="S") then msoma=msoma+10
    if (msp4="S") then msoma=msoma+10
    if (msp5="S") then msoma=msoma+10

%>
<!--#include file="cnxSQLServer.asp"-->
<% 
	on error resume next
	err.number=0

	'incluir o registro
	MyConn.Begintrans

	'vou incluir a caixa
	if err.number=0 then
   	    cSQL = "INSERT INTO tb_check(id_protocolo, sf2, sf3, sf4, sf5, so1, so2, so3, so4, so5, sp1, sp2, sp3, sp4, sp5, soma) VALUES("
        cSQL = cSQL & mid_protocolo & ","
        cSQL = cSQL & "'" & msf2 & "',"
        cSQL = cSQL & "'" & msf3 & "',"
        cSQL = cSQL & "'" & msf4 & "',"
        cSQL = cSQL & "'" & msf5 & "',"
        cSQL = cSQL & "'" & mso1 & "',"
        cSQL = cSQL & "'" & mso2 & "',"
        cSQL = cSQL & "'" & mso3 & "',"
        cSQL = cSQL & "'" & mso4 & "',"
        cSQL = cSQL & "'" & mso5 & "',"
        cSQL = cSQL & "'" & msp1 & "',"
        cSQL = cSQL & "'" & msp2 & "',"
        cSQL = cSQL & "'" & msp3 & "',"
        cSQL = cSQL & "'" & msp4 & "',"
        cSQL = cSQL & "'" & msp5 & "',"
        cSQL = cSQL & msoma & ""
   	    cSQL = cSQL & ")"
        MyConn.Execute(cSQL)
    end if  

    'vou gerar os telefones	
	if err.number=0 then
	   vetor = split(mfone," ")
       for each item in vetor
           response.write(x & "<br />")

           if len(trim(item))>0 then
                cSQL = "INSERT INTO tb_check_sf1(id_protocolo, fone) VALUES("
                cSQL = cSQL & mid_protocolo & ","
                cSQL = cSQL & "'" & mid(item,1,20) & "')"
                MyConn.Execute(cSQL)     
           end if     
       next
  	end if

	if err.number=0 then
	   MyConn.CommitTrans
		   
       MyConn.close
  	   set MyConn=nothing	

       Session("gvSoma")=msoma
		   
	   'vou para tela do checklist
	   Response.Redirect "checkC.asp?ret=0"
	else
       MyConn.RollbackTrans
    
       MyConn.close
       set MyConn=nothing				

  	   'volto a tela pai
	   Response.Redirect "checkA.asp?ret=Inclusão Invalida!" & abs(err.number) & "-" & replace(err.description,"'"," ")
	end if
%>
