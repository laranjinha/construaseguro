<%
	if Session("LoginAtivo") <> 1 Then
		Response.Redirect "logA.asp?ret=0"
	else
		Response.Redirect "portal.asp"
	end if
%>