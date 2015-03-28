<%
 
       xemail =request.querystring("email")
 
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("galeno.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr

        sql = " SELECT * " & _
              " FROM registro" & _
              " WHERE eMail = '" & xemail & "'"
			  
           '			  response.Write "SQL--------->" & sql & "<---------" 
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, cn, 3, 3 

         if Not rs.EOF Then 
         xnombre = rs.Fields("Nombre")
		 xapellido = rs.Fields("Apellido")
		 response.Write xNombre & " " & xapellido
		 session("xemail") = xemail
 
		 End if
					    
%>
<HTML>
<HEAD>
<TITLE>SIstema INterconectado de Recepción e Envío de Datos</TITLE>
<style type="text/css">
<!--
.style1 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12;
	color: #000080;
}
.style3 {font-family: Verdana, Arial, Helvetica, sans-serif}
.style4 {font-size: 12}
.style5 {color: #000080}
.style6 {font-family: Verdana, Arial, Helvetica, sans-serif; color: #000080; }
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
<BODY Background="back.jpg">
<div align="left"></div>

<FORM METHOD="Post" name="Login" ACTION="update.asp">
    
  <div align="center">    <TABLE border=0 id=TABLE1 width="427">
      <TBODY>
        <TR>
          <TD colSpan=2><div align="left">
            
</div></TD>
          <TD colSpan=2>&nbsp;</TD>
        </TR>
        <TR>
          <TD colSpan=2><div align="left"><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Especialidad:</B></FONT><BR>
              <INPUT size=27 
                    name=txt_Especialidad>
          </div></TD>
          <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Sub-especialidad:</B></FONT><BR>
              <INPUT size=27 
                    name=txt_Subespecialidad></TD>
        </TR>
        <TR>
          <TD colSpan=2><div align="left"><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Celular</B></FONT><BR>
              <INPUT size=27 
                    name=txt_Celular>
          </div></TD>
          <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Email:</B></FONT><br>
              <INPUT size=27 name=txt_email></TD>
        <TR>
          <TD colSpan=2><div align="left"><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Fecha de Cumplea&ntilde;os:</B></FONT><BR>
              <INPUT size=27 
                    name=txt_FechaCumpleanos Value="<%=FormatDateTime(Date,2)%>">
          <A href="javascript:cal1.popup();"><IMG height=16 alt="Click aqui para seleccionar una Fecha..." src  ="img/cal.gif" width=16 border=0 ></A> </div></TD>
        </TR>
        <TR>
          <TD width=90><div align="left"></div></TD>
          <TD width=161></TD>
          <TD width=162><BR></TD>
        </TR>
        <TR>
          <TD colSpan=2><div align="left"><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1>Hospital donde trabaja</FONT><BR>
              <INPUT size=27 name=txt_HospitalTrabaja>
          </div></TD>
          <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1>Cargo</FONT><BR>
              <INPUT size=27 name=txt_Cargo></TD>
        </TR>
        <TR>
          <TD colSpan=2><div align="left"><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1>Clinica donde Trabaja</FONT><BR>
              <INPUT size=27 name=txt_ClinicaTrabaja>
          </div></TD>
          <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1>Consultorio</FONT><BR>
              <INPUT size=27 name=txt_Consultorio></TD>
        </TR>
        <tr>
          <TD colSpan=2><div align="left"><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1>Tel&eacute;fono</FONT><BR>
              <INPUT size=27 name=txt_telefono>
          </div></TD>
          <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1>Fax</FONT><BR>
              <INPUT size=27 name=txt_Fax></TD>
        </tr>
        <TR>
          <TD colSpan=4 height="51"><div align="left"><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Temas de Inter&eacute;s:</B></FONT><BR>
              <TEXTAREA name=txt_TemasInteres cols=50 rows="2"></TEXTAREA>
          </div></TD>
        </TR>
        <TR><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> &nbsp;</font> </strong></TR>
        <TR>
          <TD colSpan=2>
            <p align="left"><B><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1>&iquest;Desea&nbsp;recibir informaci&oacute;n sobre&nbsp;nuestros Productos?&nbsp;&nbsp;<BR>
          Si
          <INPUT type=radio CHECKED value=SI name=txt_SINO>
          No<FONT size=2>
          <INPUT type=radio value=NO name=txt_SINO>
&nbsp;</FONT></font></b></p>            <p align="left"></P>            <P align=left><b><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1></font>&nbsp;&nbsp;
    <input type="submit" name="Submit" value="Actualizar">
</b></P></TD>
        </TR>
      </TBODY>
    </TABLE>
    <p></P>
    <p></p>
  </div>
    
</FORM>
<form>
<div align="center">
<p><input type="button" value="Volver" onclick="history.go(-1)"></p>
</div>
</form>
</BODY>
</HTML>
