<!--#include file="funcoes.asp"-->
<!--#include file="cnxSQLServer.asp"-->
<%
	on error resume next 

    'Pego os dados
	mUsuario=request.form("txtlog")
	mSenha=request.form("txtsen")
    mAchou=0

    'Evitar 
	mUsuario = EvitarPwdSql(mUsuario)
	mSenha = EvitarPwdSql(mSenha)

   	err.number=0
   
    cSQL="SELECT * FROM tb_usuarios WHERE usuario='" & mUsuario & "' and senha='" & mSenha & "' and ativo='S'"
    Set oRs = MyConn.Execute(cSQL)

   	if oRs.eof then
		Session("gvUsuario")=""
        Session("gvUsuarioID")=""
      	Session("gvNomeUsuario")=""
      	Session("gvPerfil")=""
        mAchou=0
   	else 
        'pego as informações
        mUsuario2=oRs("usuario")
		mSenha2=oRs("senha")

        'comparo exatamente as string - tem que estar exatamente iguais maiusculas e minusculas
        if strcomp(mUsuario,mUsuario2)=0 and strcomp(mSenha,mSenha2)=0 then    
      	   Session("gvUsuario")=oRs("usuario")
           Session("gvUsuarioID")=oRs("id_usuario")
      	   Session("gvNomeUsuario")=oRs("nome")
      	   Session("gvPerfil")=oRs("perfil")
           Session("gvSistemas")="PJ,PF,"  
           mAchou=1
        else
           mAchou=0
        end if
   	end if   

    oRs.close
    Set oRs=nothing
   
    MyConn.close
    Set MyConn=nothing

    if err.number>0 or mAchou=0 then
       if mAchou<>0 then
          Response.Redirect "logA.asp?ret=" & err.number & " - " & err.description
       else
          Response.Redirect "logA.asp?ret=Login Invalido! Lembre-se de digitar exatamente igual pois letras maiusculas e minusculas sao diferentes" 
       end if	  
    else
      Session("LoginAtivo") = 1
      Response.Redirect "portal.asp"
    end if
%>

