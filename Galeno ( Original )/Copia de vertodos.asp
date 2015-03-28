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
.style9 {color: #000099}
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
<BODY BGCOLOR=FFFFFF>
<div id="Layer2" style="position:absolute; left:283px; top:182px; width:159px; height:34px; z-index:2"><form>
<div align="center">
<p><input type="button" value="Volver" onclick="history.go(-1)"></p>
</div>
</form></div>
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

<H4 align="center"><font color="#000080" face="Verdana, Arial, Helvetica, sans-serif"><!-- <% response.Write(session("Clinica")) %> --></font></h4>
<div id="Layer1" "style="position:absolute; width:845px; height:98px; z-index:1; left: 13px; top: 82px; background-color: #000000; layer-background-color: #000000; border: 1px none #000000; visibility: visible; overflow: scroll;"> 
  <table width="100%" border="0">
   <tr bgcolor="#000099"> 
    <td colspan="10">
	<p align="center"><b><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;Listado 
      de Usuarios Registrados</font></b></td>
  </tr>
  <tr bgcolor="#666666"> 
    <td><b><i><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Correo Electrónico</font></i></b></td>
    <td><b><i><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Nombre</font></i></b></td>
    <td><b><i><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Apellido</font></i></b></td>
    <td><b><i><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Especialidad</font></i></b></td>
    <td><b><i><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Subespecialidad</font></i></b></td>
    <td><b><i><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Celular</font></i></b></td>
	 <td><b><i><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefono</font></i></b></td>
	  <td><b><i><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></i></b></td>
	   <td><b><i><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></i></b></td>
  </tr>
  <tr>
  
    <% if Not rs.EOF then %>
    <% Do while Not rs.EOF %>   
    <tr bgcolor="#CCCCCC"> 
    <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="mailto:<%=rs("email")%>"><%=rs("email")%></a></font></td>
	<td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="datosindividuales.asp?idcliente=<%=rs("email")%>"><%=rs("Nombre")%></a></font></td>
    <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="datosindividuales.asp?idcliente=<%=rs("email")%>"><%=rs("Apellido")%></font></td>
    <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("Especialidad")%></font></td>
    <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("Subespecialidad")%></font></td>
	<td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("Celular")%></font></td>
	<td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("Telefono")%></font></td>
	<td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("Fax")%></font></td>
    <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="eliminar.asp?borrar=true&idcliente=<%=rs("email")%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" style="text-decoration:none"><b>Borrar</b></font></a></font></td>
    </tr>
    <% rs.MoveNext
	   Loop
	 %>
    <% Else 
	    response.Write("No se han encontrado casos")
	 end if
	 %>
  </table>
</div>

</BODY>
</HTML>
