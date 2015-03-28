<% if request.form("Comportamiento") = "true" then

    xNombre= request.Form("txt_Nombre")
    xApellido= request.Form("txt_Apellido") 
    xEspecialidad= request.Form("txt_Especialidad") 
    xSubEspecialidad= request.Form("txt_SubEspecialidad")
	xconsultorio= request.Form("txt_consultorio")
	xHospitalTrabaja= request.Form("txt_HospitalTrabaja")
	xcargo= request.Form("txt_cargo")
	xclinicatrabaja = request.Form("txt_clinicatrabaja")
	xTelefono1= request.Form("txt_Telefono1")
	xFax1= request.Form("txt_Fax1")
	xConsultorio2= request.Form("txt_Consultorio2")
	xPiso2= request.Form("txt_Piso2")
	xTelefono= request.Form("txt_Telefono")
    xFax= request.Form("txt_Fax")
    xCelular= request.Form("txt_Celular")
	xemail= request.Form("txt_email")
	xFechaCumpleanos= request.Form("txt_FechaCumpleanos")
	xdondereside= request.Form("txt_dondereside")
	xTemasInteres= request.Form("txt_TemasInteres")
	xVariable= request.Form("txt_SINO")


        openstr = "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("galeno.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
        
    

    
      sqlvalida = " SELECT * " & _
              " FROM registro" & _
              " WHERE eMail = '" & xemail & "'"

     


      Set rs = Server.CreateObject("ADODB.Recordset")
      rs.Open sqlvalida, cn, 3, 3 

      
       

      if rs.eof then
                     
         		      
		    sql = ""
			Sql  = "Insert Into Registro"                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
			sql = sql & " ( "
			Sql = Sql & " Nombre,"
			Sql = Sql & " Apellido,"
			Sql = Sql & " Especialidad,"
			Sql = Sql & " SubEspecialidad,"
			Sql = Sql & " consultorio,"
			Sql = Sql & " Cargo,"
			Sql = Sql & " ClinicaTrabaja,"
			Sql = Sql & " HospitalTrabaja,"
			Sql = Sql & " Telefono,"
			Sql = Sql & " Fax,"
			Sql = Sql & " Celular,"
			Sql = Sql & " email,"
			Sql = Sql & " RecibirInformacion,"
			Sql = Sql & " FechaCumpleanos,"
			Sql = Sql & " DondeReside,"
			Sql = Sql & " TemasInteres"
				
			Sql = Sql & " ) "
			Sql = Sql & " Values "
			Sql = Sql & " ("
		
			Sql = Sql & "'" & xNombre & "',"
		 	Sql = Sql & "'" & xApellido & "',"
			Sql = Sql & "'" & xEspecialidad & "',"
		 	Sql = Sql & "'" & xSubEspecialidad & "',"
			Sql = Sql & "'" & xconsultorio & "',"
			Sql = Sql & "'" & xcargo & "',"
			Sql = Sql & "'" & xClinicaTrabaja & "',"	
			Sql = Sql & "'" & xHospitalTrabaja & "',"		
			Sql = Sql & "'" & xTelefono & "',"
    	    Sql = Sql & "'" & xFax & "',"
			Sql = Sql & "'" & xCelular & "',"
			Sql = Sql & "'" & xemail & "',"
			Sql = Sql & "'" & xVariable & "',"
			Sql = Sql & "'" & xFechaCumpleanos & "',"
			Sql = Sql & "'" & xDondeReside & "',"
		   	Sql = Sql & "'" & xTemasInteres & "'"
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
<!-- INICIO codigo contadores.miarroba.com -->
<SCRIPT LANGUAGE="JavaScript" src="http://contadores.miarroba.com/ver.php?id=428382"></SCRIPT>
<!-- FIN codigo contadores.miarroba.com -->
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
//-->
</script>
<!-- American format mm/dd/yyyy -->
<script language="JavaScript" src="calendar2.js"></script><!-- Date only with year scrolling -->
<!-- European format dd-mm-yyyy -->
<script language="JavaScript" src="calendar1.js"></script><!-- Date only with year scrolling -->
</head>

<body bgcolor="#339966" leftmargin="50" topmargin="10" rightmargin="50">
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
        <td height="629" valign="baseline"><p><img src="img/banner-guia.gif" width="750" height="63"></p>
            <FORM action="boletin.asp" method="post" name="frmReg" onSubmit="MM_validateForm('txt_Nombre','','R','txt_Apellido','','R','txt_Especialidad','','R','txt_email','','RisEmail');return document.MM_returnValue">
                 
  <font color="#000080">
     
  <INPUT type="hidden" value="true" name="Comportamiento">
  </font><font color="#000080">
  <BR>
  </font><FONT face="Courier New"><B>Introduzca aquí sus datos si desea recibir nuestro Boletín Informativo:</B></FONT><TABLE border=0 id=TABLE1 width="434">
    <TBODY>
      <TR>
        <TD width="254">
		<B>
		<FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#FF0000 
                    size=1><span style="background-color: #FFFFFF">*</span></FONT><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#000080 
                    size=1> </FONT></B><font face="Courier New">Nombre:</font><font color="#000080"><BR>
            <INPUT name=txt_Nombre size=27></font></TD>
        <TD width="170">
		<B>
		<FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#FF0000 
                    size=1><span style="background-color: #FFFFFF">* </span></FONT></B>
		<font face="Courier New">Apellido:</font><font color="#000080"><BR>
            <INPUT size=27 
                  name=txt_Apellido></font></TD>
      </TR>
      <TR>
        <TD width="254">
		<b>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FF0000">
		<span style="background-color: #FFFFFF">
		*</span></font></b><font face="Courier New">
		Especialidad</font><font color="#000080"><BR>
            <Select name=txt_Especialidad>
  			<%
  			   			 
  			 openstr = "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("galeno.mdb")
        		Set cn = Server.CreateObject("ADODB.Connection")
              cn.Open openstr
        

  			  SQLCombo = "SELECT * FROM especialidadesmedicas"
			  Set rsCombo = cn.Execute(SQLCombo)
			  While Not rsCombo.EOF
			  %>
			  <option value="<%=rsCombo("especialidades")%>"><%=rsCombo("especialidades")%></option>
			  <%
			  rsCombo.MoveNext
			  Wend
			  rsCombo.Close
			  %>
	  		</Select></font></TD>
        <TD width="170">
		<font face="Courier New">Sub-especialidad:</font><font color="#000080"><BR>
            <INPUT size=27 
                    name=txt_Subespecialidad></font></TD>
      </TR>
      <TR>
       <TD width="254">
		<B>
		<FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#FF0000 
                    size=1>*</FONT><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#000080 
                    size=1> </FONT></B><font face="Courier New">Celular</font><font color="#000080"><BR>
            <INPUT size=27 
                    name=txt_Celular></font></TD>
       <TD width="170">
		<b>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FF0000">
		* </font></b>
		<font face="Courier New">Email:</font><font color="#000080"><br>
            <INPUT size=27 name=txt_email></font></TD><TR>
          <TD width="254" height="40">
			<font face="Courier New">Fecha de Cumpleaños:</font><font color="#000080"><BR>
            <INPUT size=27 
                    name=txt_FechaCumpleanos Value="<%=FormatDateTime(Date,2)%>">	
					</font>	
					<A href="javascript:cal1.popup();"><font color="#000080"><IMG height=16 alt="Click aqui para seleccionar una Fecha..." src  ="img/cal.gif" width=16 border=0 ></font></A><font color="#000080">
			</font>
</TD>
      </TR>
  <TR>      
        <TD colspan="2">
		<B>
		<FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#FF0000 
                    size=1><span style="background-color: #FFFFFF">&nbsp;</span>*</FONT></B><font face="Courier New"> Donde reside :</font><font color="#000080">
		</font>
		<FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B><font color="#000080"><input type="text" name="txt_dondereside" size="66"></font></B></FONT></TD>
        </TR>
      <TR>
        <TD width="254" height="42"><font face="Courier New">Hospital donde trabaja</font><font color="#000080"><b><BR>
            </b>
            <INPUT size=22 name=txt_HospitalTrabaja style="font-weight: 700"></font></TD>
        <TD width="170" height="42">
		<font face="Courier New">Cargo</font><font color="#000080">
            &nbsp;</b><INPUT size=27 
                    name=txt_Cargo style="font-weight: 700"></font></TR>
      <TR>
        <TD width="254"><font face="Courier New">Clinica donde Trabaja</font><font color="#000080"><b><BR>
            </b>
            <INPUT size=22 name=txt_ClinicaTrabaja style="font-weight: 700"></font></TD>
        <TD width="170"><font face="Courier New">Consultorio</font><font color="#000080"><b><BR>
            </b>
            <INPUT size=27 name=txt_Consultorio style="font-weight: 700"></font></TD>
      </TR>
       <tr><TD width="254"><font face="Courier New">Teléfono</font><font color="#000080"><b><BR>
            </b>
            <INPUT size=22 name=txt_telefono style="font-weight: 700"></font></TD>
           <TD width="170"><font face="Courier New">Fax</font><font color="#000080"><b><BR>
            </b>
            <INPUT size=27 name=txt_Fax style="font-weight: 700"></font></TD>  </tr>
      <TR>
        <TD colSpan=2 height="51">
		<font face="Courier New">Temas de Interés:</font><font color="#000080"><BR>
		<TEXTAREA name=txt_TemasInteres cols=50 rows="2"></TEXTAREA></font></TD>
      </TR>
      <TR><font color="#000080"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      &nbsp;</font> </strong></font></TR> 
      <TR>
         <TD width="254">
			<p align="center"><font face="Courier New">¿Desea&nbsp;recibir información sobre&nbsp;nuestros 
      Productos?&nbsp;&nbsp;</font><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#000080 
                    size=1><B><BR> </b></FONT><font face="Courier New">Si</font><B><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><INPUT type=radio CHECKED value=SI name=txt_SINO></FONT></b><font face="Courier New"> 
      No</font><B><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><FONT size=2><font color="#000080"><INPUT type=radio value=NO name=txt_SINO></font><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#000080 
                    size=1>&nbsp;</FONT></FONT><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#000080 
                    size=1></p>
			</FONT>
     
			</P>
      		</FONT>
     
      <P align=center><font color="#000080">&nbsp;&nbsp;
        <input type="submit" name="Submit" value="Registrarse"></font></P></TD>
     
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
				               <script language="JavaScript">
		<!-- // create calendar object(s) just after form tag closed
			 // specify form element as the only parameter (document.forms['formname'].elements['inputname']);
			 // note: you can have as many calendar objects as you need for your application
			var cal1 = new calendar2(document.forms['frmReg'].elements['txt_FechaCumpleanos']);
			cal1.year_scroll = true;
			cal1.time_comp = false;
		//-->
		</script>
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
