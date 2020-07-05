<%
	dim mMenu

	'Verifico o nivel do menu
	if Session("gvMenu")="2" then
	   m_dir="../"
	else
	   m_dir=""
	end if  
%>
		<header>
			<div id="cab-linha"> </div>
			<div id="cab-centro">			
				<h1 class="logoleft">
					<a href="portal.asp" title="Construa Seguro / Sebrae"></a>
				</h1>   
				<p class="titulo">Construa Seguro / Sebrae<p>
			</div>
		</header> 

		<div id="menubarra">
			<%if Session("LoginAtivo")=1 then%>		
				<ul id="menu_dropdown" class="menubar">
					<li class="submenu"><a href="sistemas.asp" class="menubold">Cadastros</a></li>  
	 				<li class="submenu"><a href="<%=m_dir%>logout.asp" class="menubold">Logout</a></li>
	   				<li class="submenu"><a href="<%=m_dir%>sair.asp" target="_self" class="menubold">Sair</a></li>
				</ul>	 
				<p class="usuario">Usuário:&nbsp;<%=Session("gvUsuario")%></p>
			<%end if%>
		</div>
