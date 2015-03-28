

<HTML>
<HEAD><TITLE>GALENOS. Modulo de Administración</TITLE>
<style type="text/css">
<!--
.style1 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12;
	color: #000080;
}
.style5 {color: #000080}
-->
</style>
</HEAD>
<BODY Background="back2.jpg" vlink="black" link="black">
<FORM METHOD="Post" name="buscar" ACTION="buscar.asp">
  <div align="center"> 
    <p align="left"><strong> 
      <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif">
	  <% response.Write "Bienvenido"  %>
      <% response.Write(session("Usuario"))  %>
      </font>
      <input type="hidden" name="consulta2" value="true">
      </strong><strong><font color="#000080" face="Verdana, Arial, Helvetica, sans-serif"> 
      
      </font></strong> </p>
  </div>
  <!--  <H5 align="center" class="style1"><strong>AUTORIZACI&Oacute;N DE SERVICIO</strong></H5>-->
  <div align="center"> 
    <div id="Layer1" style="position:absolute; width:490px; height:97px; z-index:1; left: 130px; top: 31px;"><strong><font size="2">      </font> </strong> 
      <table width="103%" height="80" border=0>
        <tr> 
          <td width="309" height="24" class="style5"><div align="center"><strong></strong>
                <span class="style1"><strong><img src="img/buscar.bmp"></strong></span> <strong><font size="2">Introduzca la direcci&oacute;n de Correo Electr&oacute;nico</font></strong></div>
          <td width="186" class="style1"><div align="center"><strong><strong>
          <strong>
          <input name="email" class="style2" size="30">
          </strong> </strong> </strong> 
          </div>
      </table>
      <input type="submit" name="Submit" value="Enviar">
    </div>
    <H4 align="center" class="style1">&nbsp;</H4>
    <div align="center"></div>
  </div>
</FORM>
<form>
<div align="center">
<p><input type="button" value="Volver" onclick="history.go(-1)"></p>
</div>
</form>
<strong> 
<%

xcorreo =  request.form("email")

if request.form("consulta2") = "true" then
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("galeno.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
        
                  
         
                 
	
        sql = " SELECT * " & _
              " FROM registro" & _
              " WHERE email = '" & xcorreo & "'"
			  
		 
         Set rs = Server.CreateObject("ADODB.Recordset")
         rs.Open sql, cn, 3, 3 

         if rs.eof then
	         response.redirect("mensaje0002.asp")
         else 			 						
			 						
				    session("email")= xcorreo
				 	response.redirect("datosindividuales.asp")
			 	
			
	 end if

	 rs.Close
	 Set rs = Nothing
		
	 cn.Close
	 Set cn = Nothing

END IF
%>

</strong>
<br>
<br>
<br><br>

</BODY>
</HTML>
