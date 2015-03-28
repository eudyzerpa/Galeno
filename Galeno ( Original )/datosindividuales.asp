<HTML>
<HEAD>
<TITLE>GALENOS. Modulo de Administración</TITLE>
<style type="text/css">
<!--
.style1 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12;
	color: #000080;
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
//-->
</script>
</HEAD>
<BODY Background="back2.jpg" vlink="black" link="black">
<strong> 
<%

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
<div id="Layer1" style="position:absolute; width:340px; height:24px; z-index:7; left: 9px; top: 1px;"> 
  <table width="101%" border="0">
    <tr> 
      <td><div align="center"><strong><font color="#000099" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
          <%response.write "<STRONG>" & "<CENTER>" & "GALENOS" & "</CENTER>" & "</STRONG>" 
	 response.write "<CENTER>" & "Usuario Registrado" & "</CENTER>" 
	 response.write "<CENTER>" & "" & "</CENTER>" & "<br>"%>
          <a href="buscar.asp"><img src="img/buscar.bmp" alt="Buscar Registro" border="0"></a><a href="registro.asp"><img src="img/nuevo.bmp" alt="Agregar Registro" border="0"></a><a href="eliminar.asp?borrar=true&idcliente=<%=rs("email")%>"><img src="img/eliminar.bmp" alt="Eliminar Registro" border="0"></a><a href="actualizar.asp?email=<%=rs("email")%>"><img src="img/logout.bmp" alt="Actualizar Registro" border="0"></a></strong></div>      </tr>
  </table>
</div>
<div id="Layer2" style="position:absolute; left:47px; top:76px; width:302px; height:106px; z-index:6"> 
  <table width="100%" border="0">
    <tr> 
      <td width="47%"><strong><font color="#000099" size="1" face="Courier New">Nombre:</font></strong></td>
      <td width="53%"><strong><font color="#000099" size="1" face="Courier New"> 
        <% =rsx.Fields("Nombre")%>
        </font></strong></td>
    </tr>
    <tr> 
      <td><strong><font color="#000099" size="1" face="Courier New">Apellidos:</font></strong></td>
      <td><strong><font color="#000099" size="1" face="Courier New"> 
        <%  =rsx.Fields("Apellido")%>
        </font></strong></td>
    </tr>
    <tr> 
      <td><strong><font color="#000099" size="1" face="Courier New">Especialidad:</font></strong></td>
      <td><strong><font color="#000099" size="1" face="Courier New"> 
        <%  =rsx.Fields("Especialidad")%>
        </font></strong></td>
    </tr>
    <tr> 
      <td><strong><font color="#000099" size="1" face="Courier New">Subespecialidad:</font></strong></td>
      <td><strong><font color="#000099" size="1" face="Courier New"> 
        <%  =rsx.Fields("Subespecialidad")%>
        </font></strong></td>
    </tr>
    <tr> 
      <td><strong><font color="#000099" size="1" face="Courier New">Celular:</font></strong></td>
      <td><strong><font color="#000099" size="1" face="Courier New"> 
        <%  =rsx.Fields("Celular")%>
        </font></strong></td>
    </tr>
    <tr> 
      <td><strong><font color="#000099" size="1" face="Courier New">Correo 
        Electronico:</font></strong></td>
      <td><strong><font color="#000099" size="1" face="Courier New"> 
        <% =rsx.Fields("email")%>
        </font></strong></td>
    </tr>
    <tr>
      <td><strong><font color="#000099" size="1" face="Courier New">Fecha de Cumpleaños:</font></strong></td>
      <td><strong><font color="#000099" size="1" face="Courier New">
        <% =rsx.Fields("FechaCumpleanos")%>
      </font></strong></td>
    </tr>
    <tr>
      <td><strong><font color="#000099" size="1" face="Courier New">Hospital donde Trabaja:</font></strong></td>
      <td><strong><font color="#000099" size="1" face="Courier New">
        <% =rsx.Fields("HospitalTrabaja")%>
      </font></strong></td>
    </tr>
    <tr>
      <td><strong><font color="#000099" size="1" face="Courier New">Cargo:</font></strong></td>
      <td><strong><font color="#000099" size="1" face="Courier New">
        <% =rsx.Fields("Cargo")%>
      </font></strong></td>
    </tr>
    <tr>
      <td><strong><font color="#000099" size="1" face="Courier New">Clinica donde Trabaja:</font></strong></td>
      <td><strong><font color="#000099" size="1" face="Courier New">
        <% =rsx.Fields("ClinicaTrabaja")%>
      </font></strong></td>
    </tr>
    <tr>
      <td><strong><font color="#000099" size="1" face="Courier New">Teléfono:</font></strong></td>
      <td><strong><font color="#000099" size="1" face="Courier New">
      <% =rsx.Fields("Telefono")%>
</font></strong></td>
    </tr>
    <tr>
      <td><strong><font color="#000099" size="1" face="Courier New">fax:</font></strong></td>
      <td><strong><font color="#000099" size="1" face="Courier New">
        <% =rsx.Fields("Fax")%>
      </font></strong></td>
    </tr>
    <tr>
      <td><strong><font color="#000099" size="1" face="Courier New">Temas de Interes:</font></strong></td>
      <td><strong><font color="#000099" size="1" face="Courier New">
        <% =rsx.Fields("TemasInteres")%>
      </font></strong></td>
    </tr>
    <tr>
      <td><strong><font color="#000099" size="1" face="Courier New">Desea recibir Información:</font></strong></td>
      <td><strong><font color="#000099" size="1" face="Courier New">
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
</strong> 
<div id="Layer4" style="position:absolute; left:94px; top:410px; width:200px; height:28px; z-index:5"> 
  <strong>
  <form>
<div align="center">
<p><input type="button" value="Volver" onclick="history.go(-1)"></p>
</div>
</form></strong></div>
<p align="center">&nbsp;</p>
</BODY>
</HTML>
