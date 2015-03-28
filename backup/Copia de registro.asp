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

<html>
<head>
<title>Planilla de Suscripcion...GALENOS 2005</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/JavaScript">
<!--
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
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' El campo debe contener una dirección de correo electrónico .\n';
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' El campo debe contener un número.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' Este campo es requerido.\n'; }
  } if (errors) alert('Error en la suscripcion a Galenos:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
<!-- American format mm/dd/yyyy -->
<script language="JavaScript" src="calendar2.js"></script><!-- Date only with year scrolling -->
<!-- European format dd-mm-yyyy -->
<script language="JavaScript" src="calendar1.js"></script><!-- Date only with year scrolling -->
</head>

<body background="img/Falogo03.gif">
<FORM action="Registro.asp" method="post" name="frmReg" onSubmit="MM_validateForm('txt_Nombre','','R','txt_Apellido','','R','txt_Especialidad','','R','txt_Sub','txt_cargo','R','txt_consult','','R','txt_Hosp','','RisNum','txt_Telefono1','','R','txt_email','','RisEmail');return document.MM_returnValue">
     
  <font color="#000080">
     
  <INPUT type="hidden" value="true" name="Comportamiento">
  </font><font color="#000080">
  <BR>
  </font>
  <P>
	<FONT face="Courier New"><B>Introduzca aquí sus datos si desea recibir nuestro Boletín Informativo:</B></FONT></P>
  <TABLE border=0 id=TABLE1 width="434">
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
<form>
<div align="center">
<p><font color="#000080"><input type="button" value="Volver" onclick="history.go(-1)"></font></p>
</div>
</form>
<font color="#000080">
<script language="JavaScript">
		<!-- // create calendar object(s) just after form tag closed
			 // specify form element as the only parameter (document.forms['formname'].elements['inputname']);
			 // note: you can have as many calendar objects as you need for your application
			var cal1 = new calendar2(document.forms['frmReg'].elements['txt_FechaCumpleanos']);
			cal1.year_scroll = true;
			cal1.time_comp = false;
		//-->
		</script>
</font>
</body>
</html>