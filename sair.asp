<!--#include file="funcoes.asp"-->
<%
	on error resume next 
	
    'Expirar pagina ao enviar.
    response.buffer=true
    response.expires = 0
    response.expiresabsolute = Now() - 3
    response.addHeader "pragma","no-cache"
    response.addHeader "cache-control","private"
    Response.CacheControl = "no-cache"

	Session("LoginAtivo")=0
	Session.Contents.RemoveAll()
	Session.Abandon
	
%>
<script>
 window.close();
</script>
