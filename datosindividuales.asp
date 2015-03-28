<% if request.form("Comportamiento") = "true" then


	
	xtitulopromocion= request.Form("txt_titulopromocion")
	xpromocion= request.Form("txt_promocion")
	


        openstr = "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("galeno.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
        
    

    
      sqlvalida = " SELECT * " & _
              " FROM promocion" & _
              " WHERE Titulopromocion = '" & xtitulopromocion & "'"

     


      Set rs = Server.CreateObject("ADODB.Recordset")
      rs.Open sqlvalida, cn, 3, 3 

      
       

      if rs.eof then
                     
         		      
		    sql = ""
			Sql  = "Insert Into promocion"                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
			sql = sql & " ( "
			Sql = Sql & " Titulopromocion,"
			Sql = Sql & " promocion"
				
			Sql = Sql & " ) "
			Sql = Sql & " Values "
			Sql = Sql & " ("
		
			Sql = Sql & "'" & xtitulopromocion & "',"
			Sql = Sql & "'" & xpromocion & "'"
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
	<table background="img/Falogo03.gif" width="317" height="631"  border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
      <tr>
        <td height="629" valign="baseline"><p>&nbsp;</p>
            <FORM action="Registropromocions.asp" method="post" name="frmReg">
     
  <font color="#000080">
     
  <%
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("galeno.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
    
	
        
		
    Sql = ""
    Sql = Sql & " Select "
    Sql = Sql & " promocion.Titulopromocion,"
    Sql = Sql & " promocion.promocion"
     
        
    Sql = Sql & " From "
    Sql = Sql & " promocion"
    
   
		 
			  
		 
         Set rs = Server.CreateObject("ADODB.Recordset")
         rs.Open sql, cn, 3, 3 

        
%>

<%  'hasta aqui llega vertodos            %>

<H4 align="center"><font color="#000080" face="Verdana, Arial, Helvetica, sans-serif"><!-- <% response.Write(session("Clinica")) %> --></font></h4>
<div id="Layer1" "style="position:absolute; width:845px; height:98px; z-index:1; left: 13px; top: 82px; background-color: #000000; layer-background-color: #000000; border: 1px none #000000; visibility: visible; overflow: scroll;"> 
  &nbsp;<%

Dim xemail

    xemail =request.querystring("idcliente")


    if xemail <> "" then
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("galeno.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
                   
        sql = " SELECT * " & _
              " FROM registro " & _
              " WHERE email = '" & xemail & "'" 

        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, cn, 3, 3 

        if rs.eof then
		    response.redirect("mensaje0002.asp")
       		              	
		
		End if

            		
	 sql2 = " SELECT * " & _
                " FROM registro " & _
                " WHERE email = '" & xemail & "'" 
			  
		 Set rsx = Server.CreateObject("ADODB.Recordset")
         	 rsx.Open sql2, cn, 3, 3 

 	      
			 	 
		
		%>
</strong> 
<div id="Layer1" style="position:absolute; width:340px; height:24px; z-index:7; left: 67px; top: 180px"> <br>
  <table width="97%" border="0">
    <tr> 
      <td>&nbsp;</tr>
    <tr> 
      <td><div align="center"><font color="#000000" size="3" face="Verdana, Arial, Helvetica, sans-serif"> 
          <strong><a href="boletin.asp">
		<img src="img/nuevo.bmp" alt="Agregar Registro" border="0" width="52" height="51"></a><a href="eliminar.asp?borrar=true&idcliente=<%=rs("email")%>"><img src="img/eliminar.bmp" alt="Eliminar Registro" border="0" width="51" height="49"></a><a href="actualizar.asp?email=<%=rs("email")%>"><img src="img/logout.bmp" alt="Actualizar Registro" border="0" width="49" height="49"></a></strong><a href="listausuarios.asp"><img border="0" src="img/volver.bmp" width="49" height="45"></a></div>      </tr>
  </table>
</div>
<div id="Layer2" style="position:absolute; left:142px; top:288px; width:388px; height:106px; z-index:6"> 
  <table width="100%" border="0">
    <tr> 
      <td width="47%"><strong><font size="2" face="Courier New" color=#000000>Nombre:</font></strong></td>
      <td width="53%"><strong><font size="2" face="Courier New" color=#000000> 
        <% =rsx.Fields("Nombre")%>
        </font></strong></td>
    </tr>
    <tr> 
      <td><strong><font size="2" face="Courier New" color=#000000>Apellidos:</font></strong></td>
      <td><strong><font size="2" face="Courier New" color=#000000> 
        <%  =rsx.Fields("Apellido")%>
        </font></strong></td>
    </tr>
    <tr> 
      <td><strong><font size="2" face="Courier New" color=#000000>Especialidad:</font></strong></td>
      <td><strong><font size="2" face="Courier New" color=#000000> 
        <%  =rsx.Fields("Especialidad")%>
        </font></strong></td>
    </tr>
    <tr> 
      <td><strong><font size="2" face="Courier New" color=#000000>Subespecialidad:</font></strong></td>
      <td><strong><font size="2" face="Courier New" color=#000000> 
        <%  =rsx.Fields("Subespecialidad")%>
        </font></strong></td>
    </tr>
    <tr> 
      <td><strong><font size="2" face="Courier New" color=#000000>Celular:</font></strong></td>
      <td><strong><font size="2" face="Courier New" color=#000000> 
        <%  =rsx.Fields("Celular")%>
        </font></strong></td>
    </tr>
    <tr> 
      <td><strong><font size="2" face="Courier New" color=#000000>Correo Electronico:</font></strong></td>
      <td><strong><font size="2" face="Courier New" color=#000000> 
        <% =rsx.Fields("email")%>
        </font></strong></td>
    </tr>
    <tr>
      <td><strong><font size="2" face="Courier New" color=#000000>Fecha de Cumpleaños:</font></strong></td>
      <td><strong><font size="2" face="Courier New" color=#000000 >
        <% =rsx.Fields("FechaCumpleanos")%>
      </font></strong></td>
    </tr>
    <tr>
      <td><strong><font size="2" face="Courier New" color =#000000>Hospital donde Trabaja:</font></strong></td>
      <td><strong><font size="2" face="Courier New" color=#000000>
        <% =rsx.Fields("HospitalTrabaja")%>
      </font></strong></td>
    </tr>
    <tr>
      <td><strong><font size="2" face="Courier New" color=#000000>Cargo:</font></strong></td>
      <td><strong><font size="2" face="Courier New" color=#000000>
        <% =rsx.Fields("Cargo")%>
      </font></strong></td>
    </tr>
    <tr>
      <td><strong><font size="2" face="Courier New" color=#000000>Clinica donde Trabaja:</font></strong></td>
      <td><strong><font size="2" face="Courier New" color=#000000>
        <% =rsx.Fields("ClinicaTrabaja")%>
      </font></strong></td>
    </tr>
    <tr>
      <td><strong><font size="2" face="Courier New" color=#000000>Teléfono:</font></strong></td>
      <td><strong><font size="2" face="Courier New" color=#000000>
      <% =rsx.Fields("Telefono")%>
</font></strong></td>
    </tr>
    <tr>
      <td><strong><font size="2" face="Courier New" color=#000000>fax:</font></strong></td>
      <td><strong><font size="2" face="Courier New" color=#000000>
        <% =rsx.Fields("Fax")%>
      </font></strong></td>
    </tr>
    <tr>
      <td><strong><font size="2" face="Courier New" color=#000000>Temas de Interes:</font></strong></td>
      <td><strong><font size="2" face="Courier New" color=#000000>
        <% =rsx.Fields("TemasInteres")%>
      </font></strong></td>
    </tr>
    <tr>
      <td><strong><font size="2" face="Courier New" color=#000000>Desea recibir Información:</font></strong></td>
      <td><strong><font size="2" face="Courier New" color=#000000>
        <% =rsx.Fields("RecibirInformacion")%>
      </font></strong></td>
    </tr>
  </table>
</div>
<strong> 
<% 

		rs.Close
		Set rs = Nothing
		
		cn.Close
		Set cn = Nothing

end if
%>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
</strong> 
<div id="Layer4" style="position:absolute; left:396px; top:235px; width:200px; height:28px; z-index:5"> 
  <strong>
  <form>
<div align="center">
<p>&nbsp;</p>
</div>
</form></strong></div>
<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>

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
