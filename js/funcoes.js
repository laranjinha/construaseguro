// JavaScript Document
//controla o menu
function horizontal() {
 
   var navItems = document.getElementById("menu_dropdown").getElementsByTagName("li");
    
   for (var i=0; i< navItems.length; i++) {
      if(navItems[i].className == "submenu")
      {
         if(navItems[i].getElementsByTagName('ul')[0] != null)
         {
            navItems[i].onmouseover=function() {this.getElementsByTagName('ul')[0].style.display="block";this.style.backgroundColor = "#FF928C";}
            navItems[i].onmouseout=function() {this.getElementsByTagName('ul')[0].style.display="none";this.style.backgroundColor = "#2D6F9B";}
         }
      }
   }
 
}

function botaotopo() 
{
   window.onscroll = function()
   {
      var top = window.pageYOffset || document.documentElement.scrollTop
      if( top > 200 ) 
      {
          var botaom = document.getElementById("topo").getElementsByTagName("img");
 		  botaom[0].style.display="block";
      } 
      else
      {
          var botaom = document.getElementById("topo").getElementsByTagName("img");
 		  botaom[0].style.display="none";
      }
   }
}

//validar data de solicitação de chamado    
function check_date(data_v) {
	var expReg = /^(([0-2]\d|[3][0-1])\/([0]\d|[1][0-2])\/[1-2][0-9]\d{2})$/;
	var msgErro = 'Formato inválido de data.';
	var vdt = new Date(); //busca data do sistema
	var vdia = vdt.getDay(); //obtém dia do sistema
	var vmes = vdt.getMonth(); //obtém mês do sistema
	var vano = vdt.getYear(); //obtém ano do sistema
	
	if ((data_v.value.match(expReg)) && (data_v.value!='')){
		var dia = data_v.value.substring(0,2);
		var mes = data_v.value.substring(3,5);
		var ano = data_v.value.substring(6,10);
		if((mes==04 && dia > 30) || (mes==06 && dia > 30) || (mes==09 && dia > 30) || (mes==11 && dia > 30)){
			alert("Dia incorreto !!! O mês especificado contém no máximo 30 dias.");
			data_v.focus();
			return false;
		} else{ //1
				if(ano%4!=0 && mes==2 && dia > 28){
					alert("Data incorreta!! O mês especificado contém no máximo 28 dias.");
					data_v.focus();
					return false;
				} else{ //2
						if(ano%4==0 && mes==2 && dia > 29){
								alert("Data incorreta!! O mês especificado contém no máximo 29 dias.");
								data_v.focus();
								return false;
						} else{ //3
								if (ano < 1900) {
										alert("Data incorreta!! Ano informado menor para atendimento.");
										data_v.focus();
										return false;
								}else{ //4
									//alert ("Data correta!");
									return true;
								} //4-else
						} //3-else
				}//2-else
		}//1-else			
	} else { //5
			alert(msgErro);
			data_v.focus();
			return false;
	} //5-else
}//fim da função para validação da data do chamado

//Função: Simula Enter=Tab
function TABEnter(oEvent)
{
  var oEvent = (oEvent)? oEvent : event;
  var oTarget =(oEvent.target)? oEvent.target : oEvent.srcElement;
  if(oEvent.keyCode==13)
    oEvent.keyCode = 9;
  if(oTarget.type=="text" && oEvent.keyCode==13)
    oEvent.keyCode = 9;
  if (oTarget.type=="radio" && oEvent.keyCode==13)
    oEvent.keyCode = 9;
}

/*Função: Simula Enter=Tab - Funciona com IE/Firefox/Chrome/Opera/Safari */
/*Precisa colocar tabindex nos inputs para poder funcionar */
function TABEnterIndex(field, event) 
{
    var nTecla = event.keyCode ? event.keyCode : event.which ? event.which : event.charCode; //Serve para a maioria 
    if (nTecla == 13) 
    {
        for (i = 0; i < field.form.elements.length; i++) 
        {
            if (field.form.elements[i].tabIndex == field.tabIndex + 1) 
            {
                field.form.elements[i].focus();
                if (field.form.elements[i].type == "text") 
                {
                    field.form.elements[i].select();
                    break;
                }
            }
        }
 	}
}

//Função Identifica o browser 
function IdentificarBrowser() 
{ 
    var browserName="msie";  /*deixo setado i.e.*/
	var ua = navigator.userAgent.toLowerCase();
	if (ua.indexOf("opera") != -1) {
		browserName = "opera";
	} else if (ua.indexOf("msie") != -1) {
		browserName = "msie";
    } else if (ua.indexOf("chrome") != -1) {
		browserName = "chrome";
	} else if (ua.indexOf("safari") != -1) {
		browserName = "safari";
	} else if (ua.indexOf("mozilla") != -1) {
		if (ua.indexOf("firefox") != -1) {
			browserName = "firefox";
		} else {
			browserName = "mozilla";
		}
	}
    return browserName;
}

//Função TxtFormat: uso para formatação de campo com data, cpf etc.
function txtBoxFormat(objForm, strField, sMask, evtKeyPress) {
	var i, nCount, sValue, fldLen, mskLen,bolMask, sCod, nTecla;
    //nTecla = evtKeyPress.keyCode;  // compativel com I.E e Chrome
    nTecla = evtKeyPress.keyCode ? evtKeyPress.keyCode : evtKeyPress.which ? evtKeyPress.which : evtKeyPress.charCode; //Serve para a maioria 

	sValue = objForm[strField].value;
	// Limpa todos os caracteres de formatação que já estiverem no campo.
	sValue = sValue.toString().replace( "-", "" );
	sValue = sValue.toString().replace( "-", "" );
	sValue = sValue.toString().replace( ".", "" );
	sValue = sValue.toString().replace( ".", "" );
	sValue = sValue.toString().replace( "/", "" );
	sValue = sValue.toString().replace( "/", "" );
	sValue = sValue.toString().replace( "(", "" );
	sValue = sValue.toString().replace( "(", "" );
	sValue = sValue.toString().replace( ")", "" );
	sValue = sValue.toString().replace( ")", "" );
	sValue = sValue.toString().replace( " ", "" );
	sValue = sValue.toString().replace( " ", "" );
	sValue = sValue.toString().replace( ":", "" );
	fldLen = sValue.length;
	mskLen = sMask.length;

	i = 0;
	nCount = 0;
	sCod = "";
	mskLen = fldLen;

while (i <= mskLen) {
	bolMask = ((sMask.charAt(i) == "-") || (sMask.charAt(i) == ":") || (sMask.charAt(i) == ".") || (sMask.charAt(i) == "/"))
	bolMask = bolMask || ((sMask.charAt(i) == "(") || (sMask.charAt(i) == ")") || (sMask.charAt(i) == " "))
	if (bolMask) {
		sCod += sMask.charAt(i);
		mskLen++; 
	}
	else {
		sCod += sValue.charAt(nCount);
		nCount++;
	}
	i++;
}
objForm[strField].value = sCod;
if (nTecla != 8 && nTecla != 9) { // backspace //tab
	if (sMask.charAt(i-1) == "9") { // apenas números...
		return ((nTecla > 47) && (nTecla < 58)); } // números de 0 a 9
	else { // qualquer caracter...
		return true;
	}
} else {
	return true;
	}
}

function CheckData(pObj) { 
  var expReg = /^((0[1-9]|[12]\d)\/(0[1-9]|1[0-2])|30\/(0[13-9]|1[0-2])|31\/(0[13578]|1[02]))\/(19|20)?\d{4}$/; 
  var aRet = true; 
  if ((pObj) && (pObj.match(expReg)) && (pObj != ''))
  { 
        var dia = parseInt(pObj.substring(0,2)); 
        var mes = parseInt(pObj.substring(3,5)); 
        var ano = parseInt(pObj.substring(6,10)); 
        if ((mes == 4 || mes == 6 || mes == 9 || mes == 11) && (dia > 30))  
          aRet = false; 
        else  
          if ((ano % 4) != 0 && mes == 2 && dia > 28)  
                aRet = false; 
          else 
                if ((ano%4) == 0 && mes == 2 && dia > 29) 
                  aRet = false; 
  }
  else  
  {
        aRet = false;   
  }	
  return aRet; 
}

function validaExtensao(id) {
   var result = true;
   var n1 = id.indexOf(".csv");
   var n2 = id.indexOf(".txt");
   if (n1 == -1 && n2 == -1) { // Arquivo não permitido
      result = false; 
   }
   return result;
}

function check_hour(hora_v){
	var expReg = /^(0[0-9]|1[0-9]|2[0-3]):[0-5][0-9]$/;
	var msgErro = 'Formato inv�lido de hora !!!';

	if ((hora_v.value.match(expReg)) && (hora_v.value!=''))
	{
	   	var jhrs = (hora_v.value.substring(0,2));
	   	var jmin = (hora_v.value.substring(3,5));
	   	if ((jhrs < 00 ) || (jhrs > 23) || ( jmin < 00) ||( jmin > 59))
	   	{
			alert("Hora Invalida !!!");
			hora_v.focus();
			return false;
		}
	}
	else
	{
		alert(msgErro);
		hora_v.focus();
		return false;
	}
	return true
}

function ValidaMail(mail)
{
    var er = new RegExp(/^[A-Za-z0-9_\-\.]+@[A-Za-z0-9_\-\.]{2,}\.[A-Za-z0-9]{2,}(\.[A-Za-z0-9])?/);
    if(typeof(mail) == "string")
	{
        if(er.test(mail))
		{ 
			return true; 
		}
    }
	else if(typeof(mail) == "object")
	{
        if(er.test(mail.value))
		{ 
            return true; 
        }
    }
	else
	{
        return false;
    }
}

// replace all em uma string 
function replaceAll(str, de, para)
{
    var pos = str.indexOf(de);
    while (pos > -1){
		str = str.replace(de, para);
		pos = str.indexOf(de);
	}
    return (str);
}

function valida_cpf(cpf)
{
      var numeros, digitos, soma, i, resultado, digitos_iguais;
      digitos_iguais = 1;
      if (cpf.length < 11)
            return false;
      for (i = 0; i < cpf.length - 1; i++)
            if (cpf.charAt(i) != cpf.charAt(i + 1))
                  {
                  digitos_iguais = 0;
                  break;
                  }
      if (!digitos_iguais)
            {
            numeros = cpf.substring(0,9);
            digitos = cpf.substring(9);
            soma = 0;
            for (i = 10; i > 1; i--)
                  soma += numeros.charAt(10 - i) * i;
            resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
            if (resultado != digitos.charAt(0))
                  return false;
            numeros = cpf.substring(0,10);
            soma = 0;
            for (i = 11; i > 1; i--)
                  soma += numeros.charAt(11 - i) * i;
            resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
            if (resultado != digitos.charAt(1))
                  return false;
            return true;
            }
      else
            return false;
}

function checkFones(fones)
{
	//separo e crio o vetor
	var res = fones.split(" ");
	var erro = 0;

	for (var i = 0, l = res.length; i < l; i++) 
    {
		res[i]=res[i].trim();  //tiro espaços

		if (res[i].length>0)  //é diferente de branco
		{
			if (res[i].length>15)  //não pode ser mais que 15
			{
				erro++;
			}

			if (res[i].length<8)  //não pode ser menor que 8
			{
				erro++;
			}

			if (!isNumber(res[i]))
			{
				erro++;
			}
		}
	}

	if (erro==0)
	{
		return true;
	}
	else
	{
		return false;
	}
}

function isNumber(n) {
    return !isNaN(parseFloat(n)) && isFinite(n);
}

function valida_cnpj(cnpj)
{
      var numeros, digitos, soma, i, resultado, pos, tamanho, digitos_iguais;
      digitos_iguais = 1;
      if (cnpj.length < 14 && cnpj.length < 15)
            return false;
      for (i = 0; i < cnpj.length - 1; i++)
            if (cnpj.charAt(i) != cnpj.charAt(i + 1))
                  {
                  digitos_iguais = 0;
                  break;
                  }
      if (!digitos_iguais)
            {
            tamanho = cnpj.length - 2
            numeros = cnpj.substring(0,tamanho);
            digitos = cnpj.substring(tamanho);
            soma = 0;
            pos = tamanho - 7;
            for (i = tamanho; i >= 1; i--)
                  {
                  soma += numeros.charAt(tamanho - i) * pos--;
                  if (pos < 2)
                        pos = 9;
                  }
            resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
            if (resultado != digitos.charAt(0))
                  return false;
            tamanho = tamanho + 1;
            numeros = cnpj.substring(0,tamanho);
            soma = 0;
            pos = tamanho - 7;
            for (i = tamanho; i >= 1; i--)
                  {
                  soma += numeros.charAt(tamanho - i) * pos--;
                  if (pos < 2)
                        pos = 9;
                  }
            resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
            if (resultado != digitos.charAt(1))
                  return false;
            return true;
            }
      else
            return false;
} 


