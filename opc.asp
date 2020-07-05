<!--#include file="funcoes.asp"-->
<!--#include file="inicial.asp"-->
<!--#include file="cnxSQLServer.asp"-->
<%
 
    'Registro entrada no sistema
    Session("gvAtividadePF")="" 
    Session("gvAtividadePJ")=""
    Session("gvIDPF")=""
    Session("gvIDPJ")=""
    Session("gvProtocolo")=""
    Session("gvIDPJPF")=""

    mopc=request.QueryString("p")
    mopc=EvitarCmdSql(mopc)

    if (mopc="2e1") then  'PF
    		mopc="PF"	
    elseif(mopc="2x1") then  'PJ
    		mopc="PJ"
    else
	    response.redirect("sistemas.asp")
    end if	

    on error resume next
    err.number=0

    'vou gerar o protocolo
	  if err.number=0 then
        cSQL = "INSERT INTO tb_protocolo(tipo) VALUES("
        cSQL = cSQL & "'" & mopc & "')"
        MyConn.Execute(cSQL)
    end if  

    'Pego o id gerado 	
	  if err.number=0 then
        cSQL = "SELECT LAST_INSERT_ID() AS NewID"
        Set oRs=MyConn.Execute(cSQL)
        Session("gvProtocolo")=trim(cstr(oRs("NewID")))
        oRs.close
   	    Set oRs=nothing		
  	end if   

   if (mopc="PF") then  'PF
      Session("gvIDPJPF")="PF"
      response.redirect("pfA.asp")
   elseif(mopc="PJ") then  'PJ
      Session("gvIDPJPF")="PJ"
      response.redirect("pjA.asp")
   end if

%>
