
 <%
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("galeno.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
    
	
        
		
    Sql = ""
    Sql = Sql & " Select Top 1 "
    Sql = Sql & " promocion.Titulopromocion,"
    Sql = Sql & " promocion.promocionvalidahasta,"
    Sql = Sql & " promocion.promocion"
   
 
        
    Sql = Sql & " From "
    Sql = Sql & " promocion"
    Sql = Sql & " Order By FechaPromocion DESC"
    
   
		 
			  
		 
         Set rs = Server.CreateObject("ADODB.Recordset")
         rs.Open sql, cn, 3, 3 

        
%>


 <%
        
	
        
		
    SqlNoticia = ""
    SqlNoticia = SqlNoticia & " Select "
    SqlNoticia = SqlNoticia & " Noticia.TituloNoticia,"
    SqlNoticia = SqlNoticia & " Noticia.Noticia"
   
 
        
    SqlNoticia = SqlNoticia & " From "
    SqlNoticia = SqlNoticia & " Noticia"
    SqlNoticia = SqlNoticia & " Order By FechaNoticia DESC"
    
   
		 
			  
		 
         Set rsNoticia = Server.CreateObject("ADODB.Recordset")
         rsNoticia.Open sqlNoticia, cn, 3, 3 

        
%>

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
<script language="JavaScript">
<!--
function FP_swapImg() {//v1.0
 var doc=document,args=arguments,elm,n; doc.$imgSwaps=new Array(); for(n=2; n<args.length;
 n+=2) { elm=FP_getObjectByID(args[n]); if(elm) { doc.$imgSwaps[doc.$imgSwaps.length]=elm;
 elm.$src=elm.src; elm.src=args[n+1]; } }
}

function FP_preloadImgs() {//v1.0
 var d=document,a=arguments; if(!d.FP_imgs) d.FP_imgs=new Array();
 for(var i=0; i<a.length; i++) { d.FP_imgs[i]=new Image; d.FP_imgs[i].src=a[i]; }
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
// -->
</script>
</head>

<body bgcolor="#339966" leftmargin="50" topmargin="10" rightmargin="50" onload="FP_preloadImgs(/*url*/'buttonC.jpg', /*url*/'buttonD.jpg')">
<div id="Layer3" style="position:absolute; left:99px; top:11px; width:256px; height:97px; z-index:3"><img src="img/logotransparent.gif" width="359" height="133"></div>
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
        <td height="629" valign="baseline"><p><img src="img/banner-guia.gif" width="750" height="63"></p>
            <FORM action="Registropromocions.asp" method="post" name="frmReg">
     
  <font color="#000080">
     
 



<H4 align="center"><font color="#000080" face="Verdana, Arial, Helvetica, sans-serif"><!-- <% response.Write(session("Clinica")) %> --></font></h4>
<div id="Layer1" "style="position:absolute; width:845px; height:98px; z-index:1; left: 13px; top: 82px; background-color: #000000; layer-background-color: #000000; border: 1px none #000000; visibility: visible; overflow: scroll;"> 
  &nbsp;<table width="100%" border="0">
   <tr bgcolor="#009933"> 
    <td colspan="2">
	<p align="center"><b><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
	&nbsp;</font><font color="#FFFFFF" size="2" face="Courier New">Promoción del 
	Momento</font></b></td>
  </tr>
  <tr bgcolor="#666666"> 
    <td width="72%"><b><font face="Courier New" size="2" color="#FFFFFF"><%=rs("Titulopromocion")%></font></b></td>
            
	</font>
	
	
            
	 <td width="26%"><b><font face="Courier New" size="2" color="#FFFFFF">Promoción valida hasta:</font></b></td>
  </tr>
     
  <font color="#000080">
     
 



  <tr>
  
    <% if Not rs.EOF then %>
    <% Do while Not rs.EOF %>   
    <tr bgcolor="#CCCCCC"> 
	<td height="59" width="72%"></b><b><font size="1" color =#000000 face="Verdana, Arial, Helvetica, sans-serif"></font>
	</b><font size="1"  color=#000000 face="Verdana, Arial, Helvetica, sans-serif"><%=rs("promocion")%></font></td>
	<td height="59" width="26%"><font size="1"  color=#000000 face="Verdana, Arial, Helvetica, sans-serif"><%=rs("promocionvalidahasta")%></font></td>
    </tr>
    <% rs.MoveNext
	   Loop
	 %>
    <% Else 
	    response.Write("No se han encontrado casos")
	 end if
	 %>
  </table>
</div>            <div align="right"></div>
            <div align="right"></div>
            <div align="right"></div>
            
	</font>
	
	
            
	<div style="position: absolute; width: 269px; height: 100px; z-index: 4; left: 744px; top: 424px" id="capa1">
		<a href="listapromociones.asp">
		<img border="0" id="img2" src="buttonB.jpg" height="20" width="100" alt="Ver todas.." fp-style="fp-btn: Border Bottom 2; fp-font: Courier New; fp-font-size: 8" fp-title="Ver todas.." onmouseover="FP_swapImg(1,0,/*id*/'img2',/*url*/'buttonC.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img2',/*url*/'buttonB.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img2',/*url*/'buttonD.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img2',/*url*/'buttonC.jpg')"></a></div>
	<p>&nbsp;</p>
	
	
            
<table width="100%" border="0">
   <tr bgcolor="#009933"> 
    <td colspan="2">
	<p align="center"><b><font color="#FFFFFF" size="2" face="Courier New">Otras 
	Noticias</font></b></td>
  </tr>
     
  <font color="#000080">
     
 



  <tr bgcolor="#666666"> 
    <td width="20%"><b><font face="Courier New" size="2" color="#FFFFFF">Nota</font></b></td>
	 <td width="79%">&nbsp;</td>
  </tr>
  <tr>
  
    <% if Not rsNoticia.EOF then %>
    <% Do while Not rsNoticia.EOF %>   
    <tr bgcolor="#CCCCCC"> 
	<td height="59" width="20%"></b><b><font size="1" color =#000000 face="Verdana, Arial, Helvetica, sans-serif"><%=rsNoticia("TituloNoticia")%></font></td>
	<td height="59" width="79%"><font size="1"  color=#000000 face="Verdana, Arial, Helvetica, sans-serif"><%=rsNoticia("Noticia")%></font></td>
    </tr>
    <% rsNoticia.MoveNext
	   Loop
	 %>
    <% Else 
	    response.Write("No se han encontrado casos")
	 end if
	 %>
  </table>
            
            
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
