<% if request.form("Comportamiento") = "true" then


	
	xtitulocomentario= request.Form("txt_titulocomentario")
	xComentario= request.Form("txt_Comentario")
	


        openstr = "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("galeno.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
        
    

    
      sqlvalida = " SELECT * " & _
              " FROM Comentario" & _
              " WHERE TituloComentario = '" & xtitulocomentario & "'"

     


      Set rs = Server.CreateObject("ADODB.Recordset")
      rs.Open sqlvalida, cn, 3, 3 

      
       

      if rs.eof then
                     
         		      
		    sql = ""
			Sql  = "Insert Into Comentario"                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
			sql = sql & " ( "
			Sql = Sql & " TituloComentario,"
			Sql = Sql & " Comentario"
				
			Sql = Sql & " ) "
			Sql = Sql & " Values "
			Sql = Sql & " ("
		
			Sql = Sql & "'" & xtitulocomentario & "',"
			Sql = Sql & "'" & xComentario & "'"
			Sql = Sql & ")"
		
         	
			cn.execute Sql,	raffected
			 
		cn.Close
	        Set cn = Nothing

						         
         if raffected > 0 then
              response.Redirect("mensaje.asp")
         else
           response.Redirect("mensaje0002.asp")

         end if
Else
response.Redirect("mensaje0001.asp")


		End if
End if

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
</head>

<body bgcolor="#339966" leftmargin="50" topmargin="10" rightmargin="50">
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
            <FORM action="RegistroComentarios.asp" method="post" name="frmReg">
     
  <font color="#000080">
     
  <%
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("galeno.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
    
	
        
		
    Sql = ""
    Sql = Sql & " Select "
    Sql = Sql & " Comentario.TituloComentario,"
    Sql = Sql & " Comentario.Comentario"
     
        
    Sql = Sql & " From "
    Sql = Sql & " Comentario"
    
   
		 
			  
		 
         Set rs = Server.CreateObject("ADODB.Recordset")
         rs.Open sql, cn, 3, 3 

        
%>

<%  'hasta aqui llega vertodos            %>

<H4 align="center"><font color="#000080" face="Verdana, Arial, Helvetica, sans-serif"><!-- <% response.Write(session("Clinica")) %> --></font></h4>
<div id="Layer1" "style="position:absolute; width:845px; height:98px; z-index:1; left: 13px; top: 82px; background-color: #000000; layer-background-color: #000000; border: 1px none #000000; visibility: visible; overflow: scroll;"> 
  &nbsp;<table width="100%" border="0">
   <tr bgcolor="#009933"> 
    <td colspan="2">
	<p align="center"><b><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
	&nbsp;</font><font face="Courier New" size="2" color="#FFFFFF">Lista de 
	Comentariso</font></b></td>
  </tr>
  <tr bgcolor="#666666"> 
    <td width="20%"><b><font face="Courier New" size="2" color="#FFFFFF">Titulo</font></b></td>
	 <td width="79%"><b><font face="Courier New" size="2" color="#FFFFFF">
		Comentario</font></b></td>
  </tr>
  <tr>
  
    <% if Not rs.EOF then %>
    <% Do while Not rs.EOF %>   
    <tr bgcolor="#CCCCCC"> 
	<td height="59" width="20%"></b><b><font size="1" color =#000000 face="Verdana, Arial, Helvetica, sans-serif"><%=rs("TituloComentario")%></font></td>
	<td height="59" width="79%"><font size="1"  color=#000000 face="Verdana, Arial, Helvetica, sans-serif"><%=rs("Comentario")%></font></td>
    </tr>
    <% rs.MoveNext
	   Loop
	 %>
    <% Else 
	    response.Write("No se han encontrado casos")
	 end if
	 %>
  </table>
</div>            <div align="center"></div>
            <div align="center"></div>
            <div align="center"></div>
            <p>&nbsp;</p>
            
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
