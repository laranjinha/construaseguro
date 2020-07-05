<%
Function Cryptografar(Senha) 
    'futuramente será implementado
    Cryptografar = Senha
End Function

function EvitarCmdSql(str)
   on error resume next
   
   str = trim(str)
   str = replace(str,"javascript","")
   str = replace(str,"JavaScript","")
   str = replace(str,"JAVASCRIPT","")
   str = replace(str,"vbscript","")
   str = replace(str,"VbScript","")
   str = replace(str,"VBSCRIPT","")
   str = replace(str,"script","")
   str = replace(str,"Script","")
   str = replace(str,"SCRIPT","")
   str = replace(str,"'or'1'='1'","")
   str = replace(str,"'OR'1'='1'","")
   str = replace(str,"'Or'1'='1'","")
   str = replace(str,"'oR'1'='1'","")
   str = replace(str,"=","")
   str = replace(str,"'","")
   str = replace(str,"""","")
   str = replace(str,chr(34),"")
   'str = replace(str,"(","")
   'str = replace(str,")","")
   str = replace(str,"<","")
   str = replace(str,">","")
   str = replace(str,"--","")
   str = replace(str,"'","")
   'str = replace(str,"#","")
   'str = replace(str,"$","")
   'str = replace(str,"%","")
   str = replace(str,"�","")
   str = replace(str,"&","")
   str = replace(str,"--","")
   str = replace(str,"*","")
   
   'retiro strings
   do while instr(lcase(str),"insert ")>0 
      str = mid(str,1,instr(lcase(str),"insert ")-1) & mid(str,instr(lcase(str),"insert ")+7)
   loop
   
   do while instr(lcase(str),"drop ")>0 
      str = mid(str,1,instr(lcase(str),"drop ")-1) & mid(str,instr(lcase(str),"drop ")+5)
   loop
   
   do while instr(lcase(str),"delet ")>0 
      str = mid(str,1,instr(lcase(str),"delet ")-1) & mid(str,instr(lcase(str),"delet ")+6)
   loop   

   do while instr(lcase(str),"alter table")>0 
      str = mid(str,1,instr(lcase(str),"alter table")-1) & mid(str,instr(lcase(str),"alter table")+11)
   loop   
   
   do while instr(lcase(str),"xp_")>0 
      str = mid(str,1,instr(lcase(str),"xp_")-1) & mid(str,instr(lcase(str),"xp_")+3)
   loop   
   
   do while instr(lcase(str),"select ")>0 
      str = mid(str,1,instr(lcase(str),"select ")-1) & mid(str,instr(lcase(str),"select ")+7)
   loop
   
   do while instr(lcase(str)," inner ")>0 
      str = mid(str,1,instr(lcase(str)," inner ")-1) & mid(str,instr(lcase(str)," inner ")+7)
   loop
   
   do while instr(lcase(str)," where ")>0 
      str = mid(str,1,instr(lcase(str)," where ")-1) & mid(str,instr(lcase(str)," where ")+7)
   loop   
   
   do while instr(lcase(str)," from ")>0 
      str = mid(str,1,instr(lcase(str)," from ")-1) & mid(str,instr(lcase(str)," from ")+6)
   loop   
   
   do while instr(lcase(str)," between ")>0 
      str = mid(str,1,instr(lcase(str)," between ")-1) & mid(str,instr(lcase(str)," between ")+9)
   loop   
   
   do while instr(lcase(str)," not ")>0 
      str = mid(str,1,instr(lcase(str)," not ")-1) & mid(str,instr(lcase(str)," not ")+5)
   loop   
   
   do while instr(lcase(str)," or ")>0 
      str = mid(str,1,instr(lcase(str)," or ")-1) & mid(str,instr(lcase(str)," or ")+4)
   loop   
   
   do while instr(lcase(str)," and ")>0 
      str = mid(str,1,instr(lcase(str)," and ")-1) & mid(str,instr(lcase(str)," and ")+5)
   loop   
   
   do while instr(lcase(str)," in ")>0 
      str = mid(str,1,instr(lcase(str)," in ")-1) & mid(str,instr(lcase(str)," in ")+4)
   loop   
   
   do while instr(lcase(str),"update")>0 
      str = mid(str,1,instr(lcase(str),"update")-1) & mid(str,instr(lcase(str),"update")+6)
   loop   
   
   do while instr(lcase(str),"-shutdown")>0 
      str = mid(str,1,instr(lcase(str),"-shutdown")-1) & mid(str,instr(lcase(str),"-shutdown")+9)
   loop   
   
   do while instr(lcase(str),"shutdown")>0 
      str = mid(str,1,instr(lcase(str),"shutdown")-1) & mid(str,instr(lcase(str),"shutdown")+8)
   loop   
      
   EvitarCmdSql = str
end function

function EvitarPwdSql(str)
   on error resume next
   
   str = trim(str)
   str = replace(str,"javascript","")
   str = replace(str,"JavaScript","")
   str = replace(str,"JAVASCRIPT","")
   str = replace(str,"vbscript","")
   str = replace(str,"VbScript","")
   str = replace(str,"VBSCRIPT","")
   str = replace(str,"script","")
   str = replace(str,"Script","")
   str = replace(str,"SCRIPT","")
   str = replace(str,"'or'1'='1'","")
   str = replace(str,"'OR'1'='1'","")
   str = replace(str,"'Or'1'='1'","")
   str = replace(str,"'oR'1'='1'","")
   
   'retiro strings
   do while instr(lcase(str),"insert")>0 
      str = mid(str,1,instr(lcase(str),"insert")-1) & mid(str,instr(lcase(str),"insert")+6)
   loop
   
   do while instr(lcase(str),"drop")>0 
      str = mid(str,1,instr(lcase(str),"drop")-1) & mid(str,instr(lcase(str),"drop")+4)
   loop
   
   do while instr(lcase(str),"delet")>0 
      str = mid(str,1,instr(lcase(str),"delet")-1) & mid(str,instr(lcase(str),"delet")+5)
   loop   

   do while instr(lcase(str),"alter table")>0 
      str = mid(str,1,instr(lcase(str),"alter table")-1) & mid(str,instr(lcase(str),"alter table")+11)
   loop   
   
   do while instr(lcase(str),"xp_")>0 
      str = mid(str,1,instr(lcase(str),"xp_")-1) & mid(str,instr(lcase(str),"xp_")+3)
   loop   
   
   do while instr(lcase(str),"select")>0 
      str = mid(str,1,instr(lcase(str),"select")-1) & mid(str,instr(lcase(str),"select")+6)
   loop
   
   do while instr(lcase(str)," inner ")>0 
      str = mid(str,1,instr(lcase(str)," inner ")-1) & mid(str,instr(lcase(str)," inner ")+7)
   loop
   
   do while instr(lcase(str)," where ")>0 
      str = mid(str,1,instr(lcase(str)," where ")-1) & mid(str,instr(lcase(str)," where ")+7)
   loop   
   
   do while instr(lcase(str)," from ")>0 
      str = mid(str,1,instr(lcase(str)," from ")-1) & mid(str,instr(lcase(str)," from ")+6)
   loop   
   
   do while instr(lcase(str)," between ")>0 
      str = mid(str,1,instr(lcase(str)," between ")-1) & mid(str,instr(lcase(str)," between ")+9)
   loop   
   
   do while instr(lcase(str)," not ")>0 
      str = mid(str,1,instr(lcase(str)," not ")-1) & mid(str,instr(lcase(str)," not ")+5)
   loop   
   
   do while instr(lcase(str)," or ")>0 
      str = mid(str,1,instr(lcase(str)," or ")-1) & mid(str,instr(lcase(str)," or ")+4)
   loop   
   
   do while instr(lcase(str)," and ")>0 
      str = mid(str,1,instr(lcase(str)," and ")-1) & mid(str,instr(lcase(str)," and ")+5)
   loop   
   
   do while instr(lcase(str)," in ")>0 
      str = mid(str,1,instr(lcase(str)," in ")-1) & mid(str,instr(lcase(str)," in ")+4)
   loop   
   
   do while instr(lcase(str),"update")>0 
      str = mid(str,1,instr(lcase(str),"update")-1) & mid(str,instr(lcase(str),"update")+6)
   loop   
   
   do while instr(lcase(str),"-shutdown")>0 
      str = mid(str,1,instr(lcase(str),"-shutdown")-1) & mid(str,instr(lcase(str),"-shutdown")+9)
   loop   
   
   do while instr(lcase(str),"shutdown")>0 
      str = mid(str,1,instr(lcase(str),"shutdown")-1) & mid(str,instr(lcase(str),"shutdown")+8)
   loop   
      
   EvitarPwdSql = str
end function

Function So_Numeros_Aqui(m_val)
    Dim strtemp
	Dim x_x
	
	on error resume next
	err.number=0
	
	if len(trim(m_val))>0 then
 	   strtemp=""
	   for x_x=1 to len(trim(m_val))
	       if instr("0123456789",mid(trim(m_val),x_x,1))>0 then
		      strtemp=strtemp & mid(trim(m_val),x_x,1)
		   end if	  
	   next 
	   
       if len(strtemp)=0 or isnull(strtemp) then    
		  strtemp="0"
       end if
	else
	   strtemp="0"
	end if   
	
	So_Numeros_Aqui=strtemp
	
end function

Function retorna_so_numeros(m_val)
    Dim strtemp
	Dim x_x
	
	on error resume next
	err.number=0
	
    strtemp=""
	if len(trim(m_val))>0 then
	   for x_x=1 to len(trim(m_val))
	       if instr("0123456789",mid(trim(m_val),x_x,1))>0 then
		      strtemp=strtemp & mid(trim(m_val),x_x,1)
		   end if	  
	   next 
	end if   
	
	retorna_so_numeros=strtemp
	
end function

Function strZero(Valor,Comprimento)
   If Len(trim(Valor)) <= Comprimento Then
      strZero = String(Comprimento - Len(trim(Valor)), "0") & Valor
   Else
      strZero = Valor
   End If
End Function

Function IIf(ByVal Condicao, ByVal Valor1, ByVal Valor2)
    If Condicao Then IIf = Valor1 Else IIf = Valor2
End Function

Function valida_data(EasTextData) 
	'retornar 0 OK 1 erro
    Dim EasDia
	dim EasMes
	dim EasAno 
    Dim eas_pode_nao 
    
    'Verifico se deixou em branco
    If Len(Trim(EasTextData)) = 0 Then
       valida_data = 0
       Exit Function
    End If
    
    'Verifico se digitou os dados no padrao
    eas_pode_nao = 0
    
    'tamanho
    If Len(Trim(EasTextData)) <> 10 Then
       eas_pode_nao = 1
    End If
    
    'barras
    If Mid(Trim(EasTextData), 3, 1) = "/" And Mid(Trim(EasTextData), 6, 1) = "/" Then
    Else
       eas_pode_nao = 1
    End If
       
    'numeros
    For x_x = 1 To 10
        If InStr("0123456789", Mid(Trim(EasTextData), x_x, 1)) = 0 Then
           If x_x <> 3 And x_x <> 6 Then
              eas_pode_nao = 1
              Exit For
           End If
        End If
    Next 
    
    If eas_pode_nao=1 Then
       valida_data = 1
       Exit Function
    End If
    
    'pego os dados
    easDia = cint(Mid(EasTextData, 1, 2))
    EasMes = cint(Mid(EasTextData, 4, 2))
    EasAno = cint(Mid(EasTextData, 7, 4))
    
    'verifico o EasMes
    If EasMes <= 0 Or EasMes >= 13 Then
       eas_pode_nao = 1
    End If
    
    'verifico o EasDia
    If easDia <= 0 Or easDia >= 32 Then
       eas_pode_nao = 1
    End If
    
    'verifico fevereiro
    If EasMes = 2 And easDia >= 30 Then
      eas_pode_nao = 1
    End If
    
    'verifico o EasMeses com 30 EasDias
    If ((EasMes = 4) Or (EasMes = 6) Or (EasMes = 9) Or (EasMes = 11)) And (easDia >= 31) Then
       eas_pode_nao = 1
    End If
    
    If eas_pode_nao = 0 Then
       valida_data = 0
    Else
       valida_data = 1
    End If
        
    Exit Function
End Function

Function DataMySQL(pData)
	Dim mNovaData

    mNovaData = pData

    'Verifico se veio algo e coloco no formato YYYY-MM-DD   
    if len(trim(mNovaData))>0 then
 	   mNovaData=mid(mNovaData,7,4) & "-" & mid(mNovaData,4,2) & "-" & mid(mNovaData,1,2)
    end if

	DataMySQL=mNovaData
End Function

Function PegaDataSystema()
    pdia	   = day(date)
    pmes	   = month(date)
    pano	   = year(date)
    pData = strZero(pdia,2) & "/" & strZero(pmes,2) & "/" & strZero(pano,4) 
	
	PegaDataSystema=pData
End Function

Function Acerta_Ano(EasTextData) 
	'retornar 0 OK 1 erro
    Dim EasDia
	dim EasMes
	dim EasAno 
    
    'checo as barras
    If Mid(Trim(EasTextData), 3, 1) = "/" And Mid(Trim(EasTextData), 6, 1) = "/" Then
	    'pego os dados
	    easDia = Mid(EasTextData, 1, 2)
	    EasMes = Mid(EasTextData, 4, 2)
	    EasAno = Mid(EasTextData, 7)

        if len(trim(EasAno))=2 then      
           Acerta_Ano=easDia & "/" & EasMes & "/20" & EasAno      
        else
           Acerta_Ano=EasTextData 
        end if    
    else
        Acerta_Ano=EasTextData       
	end if  
        
    Exit Function
End Function
%>
