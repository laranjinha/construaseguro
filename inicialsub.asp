<%
   'Expirar pagina ao enviar.
   response.buffer=true
   response.expires = 0
   response.expiresabsolute = Now() - 3
   response.addHeader "pragma","no-cache"
   response.addHeader "cache-control","private"
   Response.CacheControl = "no-cache"
   
   if Session("LoginAtivo") <> 1 Then
      Response.Write "<script>"
      Response.Write "window.open('logA.asp?ret=0','_parent')"
      Response.Write "</script>"	
      Response.end  	
   end If
 
%>