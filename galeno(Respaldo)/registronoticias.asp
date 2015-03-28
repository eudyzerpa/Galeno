<% if request.form("Comportamiento") = "true" then


	
	xtitulonoticia= request.Form("txt_titulonoticia")
	xnoticia= request.Form("txt_noticia")
	


        openstr = "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("galeno.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
        
    

    
      sqlvalida = " SELECT * " & _
              " FROM noticia" & _
              " WHERE Titulonoticia = '" & xtitulonoticia & "'"

     


      Set rs = Server.CreateObject("ADODB.Recordset")
      rs.Open sqlvalida, cn, 3, 3 

      
       

      if rs.eof then
                     
         		      
		    sql = ""
			Sql  = "Insert Into noticia"                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
			sql = sql & " ( "
			Sql = Sql & " Titulonoticia,"
			Sql = Sql & " Fechanoticia,"
			Sql = Sql & " noticia"
				
			Sql = Sql & " ) "
			Sql = Sql & " Values "
			Sql = Sql & " ("
		
			Sql = Sql & "'" & xtitulonoticia & "',"
			Sql = Sql & "Now(),"
			Sql = Sql & "'" & xnoticia & "'"
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
            <FORM action="Registronoticias.asp" method="post" name="frmReg">
     
  <font color="#000080">
     
  <INPUT type="hidden" value="true" name="Comportamiento">
  </font>
  <P>
	<FONT face="Courier New"><B>Introduzca aquí la Noticia del Día:</B></FONT></P>
  <TABLE border=0 id=TABLE1 width="434">
    <TBODY>
  <TR>      
        <TD>
		<B>
		<FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#FF0000 
                    size=1><span style="background-color: #FFFFFF">&nbsp;</span>*</FONT></B><font face="Courier New"> 
		Titulo de Noticia :</font><font color="#000080">
		</font>
		<FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B><font color="#000080"><input type="text" name="txt_titulonoticia" size="66"></font></B></FONT></TD>
        </TR>
      <TR>
        <TD height="51">
		<font face="Courier New">Noticia</font><font color="#000080"><BR>
		<TEXTAREA name=txt_noticia cols=50 rows="13"></TEXTAREA></font></TD>
      </TR>
      <TR><font color="#000080"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      &nbsp;</font> </strong></font></TR> 
      <TR>
         <TD width="254">
			<p align="center"><B>
			<FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1></p>
			     
			</P>
      		</FONT>
     
      <P align=center><font color="#000080">&nbsp;&nbsp;
        <input type="submit" name="Submit" value="Enviar noticia"></font></P></TD>
     
      </TR>      
    </TBODY>
  </TABLE>
	<p align="center">&nbsp;</p>
	</P>
  <p></p>
</FORM>
            <div align="center"></div>
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
