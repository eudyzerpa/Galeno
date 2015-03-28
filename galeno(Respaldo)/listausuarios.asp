<%
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("galeno.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
    
	
        
		
    Sql = ""
    Sql = Sql & " Select "
    Sql = Sql & " Registro.nombre,"    
    Sql = Sql & " Registro.Apellido,"
    Sql = Sql & " Registro.Especialidad,"    
    Sql = Sql & " Registro.Subespecialidad,"
    Sql = Sql & " Registro.HospitalTrabaja,"
    Sql = Sql & " Registro.Cargo,"
    Sql = Sql & " Registro.ClinicaTrabaja,"
    Sql = Sql & " Registro.Consultorio,"

    Sql = Sql & " Registro.Telefono,"
    Sql = Sql & " Registro.Fax,"
    Sql = Sql & " Registro.Celular,"
    
    Sql = Sql & " Registro.email,"
    Sql = Sql & " Registro.FechaCumpleanos,"
    Sql = Sql & " Registro.TemasInteres,"
    Sql = Sql & " Registro.RecibirInformacion"
     
        
    Sql = Sql & " From "
    Sql = Sql & " registro"
    
   
		 
			  
		 
         Set rs = Server.CreateObject("ADODB.Recordset")
         rs.Open sql, cn, 3, 3 

        
%>

<%  'hasta aqui llega vertodos            %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Galeno especialidades médicas c.a</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.style1 {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 12px}
.style4 {
	font-size: 14px;
	font-weight: bold;
}
.style6 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; }
.style8 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 14px;
}
a {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12px;
	color: #FFFFFF;
}
a:hover {
	color: #333333;
}
-->
</style>
<script language="JavaScript" type="text/JavaScript">

<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//
function FP_preloadImgs() {//v1.0
 var d=document,a=arguments; if(!d.FP_imgs) d.FP_imgs=new Array();
 for(var i=0; i<a.length; i++) { d.FP_imgs[i]=new Image; d.FP_imgs[i].src=a[i]; }
}

function FP_swapImg() {//v1.0
 var doc=document,args=arguments,elm,n; doc.$imgSwaps=new Array(); for(n=2; n<args.length;
 n+=2) { elm=FP_getObjectByID(args[n]); if(elm) { doc.$imgSwaps[doc.$imgSwaps.length]=elm;
 elm.$src=elm.src; elm.src=args[n+1]; } }
}

function FP_getObjectByID(id,o) {//v1.0
 var c,el,els,f,m,n; if(!o)o=document; if(o.getElementById) el=o.getElementById(id);
 else if(o.layers) c=o.layers; else if(o.all) el=o.all[id]; if(el) return el;
 if(o.id==id || o.name==id) return o; if(o.childNodes) c=o.childNodes; if(c)
 for(n=0; n<c.length; n++) { el=FP_getObjectByID(id,c[n]); if(el) return el; }
 f=o.forms; if(f) for(n=0; n<f.length; n++) { els=f[n].elements;
 for(m=0; m<els.length; m++){ el=FP_getObjectByID(id,els[n]); if(el) return el; } }
 return null;
}
-->
</script>

</head>

<body bgcolor="#339966" leftmargin="50" topmargin="10" rightmargin="50" vlink=#339966 onload="FP_preloadImgs(/*url*/'button10.jpg',/*url*/'button11.jpg')">
<div id="Layer3" style="position:absolute; left:99px; top:11px; width:256px; height:97px; z-index:3"><img src="img/logotransparent.gif" width="359" height="133"></div>
<script language="JavaScript"> 
if(window.screen.availWidth == 1280)window.parent.document.body.style.zoom="120%" 
if(window.screen.availWidth == 1152)window.parent.document.body.style.zoom="108%" 
if(window.screen.availWidth == 1024)window.parent.document.body.style.zoom="96%" 
if(window.screen.availWidth == 800)window.parent.document.body.style.zoom="75%" 
if(window.screen.availWidth == 640)window.parent.document.body.style.zoom="60%" 
</script>
<table width="82%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="45%" height="137" valign="top" bgcolor="#FFFFFF">
      <blockquote>
        <p class="style1 style4">&nbsp;</p>
    </blockquote></td>
    <td width="55%" valign="top" bgcolor="#FFFFFF"><img src="img/top2.jpg" width="429" height="137"></td>
  </tr>
  <tr align="center">
    <td height="59" colspan="2" valign="top" bgcolor="#FFFFFF"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="764" height="58">
      <param name="movie" value="menugaleno.swf">
      <param name="quality" value="high">
      <embed src="menugaleno.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="764" height="58"></embed>
    </object></td>
  </tr>
  <tr>
    <td height="644" colspan="2" valign="top" bgcolor="#FFFFFF">
	<table background="img/Falogo03.gif" width="99" height="631"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
      <tr>
        <td height="629" valign="baseline"><p>&nbsp;</p>
            <FORM action="boletin.asp" method="post" name="frmReg" onSubmit="MM_validateForm('txt_Nombre','','R','txt_Apellido','','R','txt_Especialidad','','R','txt_email','','RisEmail');return document.MM_returnValue">
                 
  <font color="#000080">
     
  <INPUT type="hidden" value="true" name="Comportamiento">
  </font>
	<p align="center"><table width="100%" border="0">
   <tr bgcolor="#cccccc"> 
    <td colspan="9">
	<p align="center"><b>
	<font size="1" face="Verdana, Arial, Helvetica, sans-serif">
	&nbsp;</font><font size="2" face="Courier New">Listado de 
	Usuarios Registrados</font></b></td>
  </tr>
  <tr bgcolor="#cccccc"> 
    <td><b><font size="2" face="Courier New">Correo Electrónico</font></b></td>
    <td><b><font size="2" face="Courier New">Nombre</font></b></td>
    <td><b><font size="2" face="Courier New">Apellido</font></b></td>
    <td><b><font size="2" face="Courier New">Especialidad</font></b></td>
    <td><b><font size="2" face="Courier New">Subespecialidad</font></b></td>
    <td><b><font size="2" face="Courier New">Celular</font></b></td>
	 <td><b><font size="2" face="Courier New">Telefono</font></b></td>
	  <td><b><font size="2" face="Courier New">Desea Recibir Información?</font></b></td>
	   <td>&nbsp;</td>
  </tr>
  <tr>
  
    <% if Not rs.EOF then %>
    <% Do while Not rs.EOF %>   
    <tr bgcolor="#339966"> 
    <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="mailto:<%=rs("email")%>"><%=rs("email")%></a></font></td>
	<td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="datosindividuales.asp?idcliente=<%=rs("email")%>"><%=rs("Nombre")%></a></font></td>
    <td><a href="datosindividuales.asp?idcliente=<%=rs("email")%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("Apellido")%></font></td>
    <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color=#FFFFFF><%=rs("Especialidad")%></font></td>
    <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color=#FFFFFF><%=rs("Subespecialidad")%></font></td>
	<td><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color=#FFFFFF><%=rs("Celular")%></font></td>
	<td><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color=#FFFFFF><%=rs("Telefono")%></font></td>
	<td><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color=#FFFFFF><%=rs("recibirinformacion")%></font></td>
    <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="eliminar.asp?borrar=true&idcliente=<%=rs("email")%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" style="text-decoration:yes"><b>Borrar</b></font></a></font></td>
    </tr>
    <% rs.MoveNext
	   Loop
	 %>
    <% Else 
	    response.Write("No se han encontrado casos")
	 end if
	 %>
  </table>
</div></p>
	</P>
  <p></p>
            
            
            
            </FORM>
            <div align="center"></div>
            <div align="center"></div>
            <div align="center"></div>
            <p align="center"><a href="controlpanel.htm">
			<img border="0" id="img1" src="button12.jpg" height="20" width="100" alt="Panel de Control" onmouseover="FP_swapImg(1,0,/*id*/'img1',/*url*/'button10.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img1',/*url*/'button12.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img1',/*url*/'button11.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img1',/*url*/'button10.jpg')" fp-style="fp-btn: Border Bottom 2" fp-title="Panel de Control"></a></p>
            
          </table></td>
		                 <p align="center"><font color="#000080">
	
                  </font></p>
                <p></P>
                <p></p>
  </tr></td>
  </tr>

  <tr align="center" valign="baseline">
    <td height="24" colspan="2"><p class="style8"><font color="#FFFFFF">www.galeno.com.ve 
        E-mail:<u> galeno@galeno.com.ve</u></font></p>    </td>
  </tr>
</table>
</body>
</html>