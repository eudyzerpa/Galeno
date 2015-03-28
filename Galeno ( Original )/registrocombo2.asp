<%

if request.form("Comportamiento") = "true" then

    xNombre= request.Form("txt_Nombre")
    xApellido= request.Form("txt_Apellido") 
    xEspecialidad= request.Form("txt_Especialidad") 
    xSubEspecialidad= request.Form("txt_SubEspecialidad")
	xconsultorio= request.Form("txt_consultorio")
	xHospitalTrabaja= request.Form("txt_HospitalTrabaja")
	xcargo= request.Form("txt_cargo")
	xTelefono1= request.Form("txt_Telefono1")
	xFax1= request.Form("txt_Fax1")
	xConsultorio2= request.Form("txt_Consultorio2")
	xPiso2= request.Form("txt_Piso2")
	xTelefono= request.Form("txt_Telefono")
    xFax= request.Form("txt_Fax")
    xCelular= request.Form("txt_Celular")
	xemail= request.Form("txt_email")
	xFechaCumpleanos= request.Form("txt_FechaCumpleanos")
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
		    
			Sql  = "Insert Into Registro"                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
			sql = sql & " ( "
			Sql = Sql & " Nombre,"
			Sql = Sql & " Apellido,"
			Sql = Sql & " Especialidad,"
			Sql = Sql & " SubEspecialidad,"
			Sql = Sql & " consultorio,"
			Sql = Sql & " Cargo,"
			Sql = Sql & " HospitalTrabaja,"
			Sql = Sql & " Telefono,"
			Sql = Sql & " Fax,"
			Sql = Sql & " Celular,"
			Sql = Sql & " email,"
			Sql = Sql & " RecibirInformacion,"
			Sql = Sql & " FechaCumpleanos,"
			Sql = Sql & " TemasInteres"
				
			Sql = Sql & " ) "
			Sql = Sql & " Values "
			Sql = Sql & " ("
		
			Sql = Sql & "'bbbbbbbbbbbbbbbbbbbbbb',"
		 	Sql = Sql & "'bbbbbbbbbbbbbbbbbbbbbb',"
			Sql = Sql & "'" & xEspecialidad & "',"
		 	Sql = Sql & "'" & xSubEspecialidad & "',"
			Sql = Sql & "'" & xconsultorio & "',"
			Sql = Sql & "'pepetrueno',"
			Sql = Sql & "'" & xHospitalTrabaja & "',"		
			Sql = Sql & "'" & xTelefono & "',"
    	    Sql = Sql & "'" & xFax & "',"
			Sql = Sql & "'" & xCelular & "',"
			Sql = Sql & "'" & xemail & "',"
			Sql = Sql & "'" & xVariable & "',"
			Sql = Sql & "'" & xFechaCumpleanos & "',"
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

<body>
<FORM action="Registrocombobox.asp" method="post" name="frmReg" onSubmit="MM_validateForm('txt_Nombre','','R','txt_Apellido','','R','txt_Especialidad','','R','txt_Sub','','R','txt_consult','','R','txt_Hosp','','RisNum','txt_Telefono1','','R','txt_email','','RisEmail');return document.MM_returnValue">
     
  <INPUT type="hidden" value="true" name="Comportamiento">
  <BR>
  <P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
              color=#996600 size=-2><B>Introduzca aquí sus datos si desea recibir nuestro Boletín Informativo:</B></FONT></P>
<P>
  <TABLE border=0 id=TABLE1 width="427">
    <TBODY>
      <TR>
        <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Nombre:</B></FONT><BR>
            <INPUT name=txt_Nombre size=27></TD>
        <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Apellido:</B></FONT><BR>
            <INPUT size=27 
                  name=txt_Apellido></TD>
      </TR>
      <TR>
        <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Especialidad:</B></FONT><BR>
            <INPUT size=27 
                    name=txt_Especialidad></TD>
        <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Sub-especialidad:</B></FONT><BR>
            <INPUT size=27 
                    name=txt_Subespecialidad></TD>
      </TR>
      <TR>
       <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Celular</B></FONT><BR>
            <INPUT size=27 
                    name=txt_Celular></TD>
       <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Email:</B></FONT><br>
            <INPUT size=27 name=txt_email></TD><TR>
          <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Fecha de Cumpleaños:</B></FONT><BR>
            <INPUT size=27 
                    name=txt_FechaCumpleanos Value="<%=FormatDateTime(Date,2)%>">	
					<A href="javascript:cal1.popup();"><IMG height=16 alt="Click aqui para seleccionar una Fecha..." src  ="img/cal.gif" width=16 border=0 ></A>
</TD>
      </TR>
  <TR>      
        <TD width=90></TD>
        <TD width=150></TD>
        <TD width=173><BR></TD></TR>
      <TR>
        <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1>Hospital donde trabaja</FONT><BR>
            <INPUT size=27 name=txt_HospitalTrabaja></TD>
        <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1>Cargo</FONT><BR>
            <Select name=txt_cargo size="1">
  			<option value="1">Enero</option>
            <option value="2">Febrero</option>
	  		</Select>

      </TR>
      <TR>
        <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1>Clinica donde Trabaja</FONT><BR>
            <INPUT size=27 name=txt_ClinicaTrabaja></TD>
        <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1>Consultorio</FONT><BR>
            <INPUT size=27 name=txt_Consultorio></TD>
      </TR>
       <tr><TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1>Teléfono</FONT><BR>
            <INPUT size=27 name=txt_telefono></TD>
           <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1>Fax</FONT><BR>
            <INPUT size=27 name=txt_Fax></TD>  </tr>
      <TR>
        <TD colSpan=4 height="51"><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Temas de Interés:</B></FONT><BR>
		<TEXTAREA name=txt_TemasInteres cols=50 rows="2"></TEXTAREA></TD>
      </TR>
      <TR><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      &nbsp;</font> </strong></TR> 
      <TR>
         <TD colSpan=2>
			<p align="center"><B><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1>¿Desea&nbsp;recibir información sobre&nbsp;nuestros 
      Productos?&nbsp;&nbsp;<BR> Si<INPUT type=radio CHECKED value=SI name=txt_SINO> 
      No<FONT size=2><INPUT type=radio value=NO name=txt_SINO>&nbsp;</FONT></p>
			</P>
      <P align=center></FONT>
     
&nbsp;&nbsp;
        <input type="submit" name="Submit" value="Registrarse"></P></TD>
     
      </TR>      
    </TBODY>
  </TABLE></P>
  <p></p>
</FORM>
<form>
<div align="center">
<p><input type="button" value="Volver" onclick="history.go(-1)"></p>
</div>
</form>
<script language="JavaScript">
		<!-- // create calendar object(s) just after form tag closed
			 // specify form element as the only parameter (document.forms['formname'].elements['inputname']);
			 // note: you can have as many calendar objects as you need for your application
			var cal1 = new calendar2(document.forms['frmReg'].elements['txt_FechaCumpleanos']);
			cal1.year_scroll = true;
			cal1.time_comp = false;
		//-->
		</script>
</body>
</html>
