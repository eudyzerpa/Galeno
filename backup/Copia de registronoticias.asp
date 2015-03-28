<% if request.form("Comportamiento") = "true" then


	
	xTituloNoticia= request.Form("txt_TituloNoticia")
	xNoticia= request.Form("txt_Noticia")
	


        openstr = "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("galeno.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
        
    

    
      sqlvalida = " SELECT * " & _
              " FROM Noticia" & _
              " WHERE TituloNoticia = '" & xTituloNoticia & "'"

     


      Set rs = Server.CreateObject("ADODB.Recordset")
      rs.Open sqlvalida, cn, 3, 3 

      
       

      if rs.eof then
                     
         		      
		    sql = ""
			Sql  = "Insert Into Noticia"                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
			sql = sql & " ( "
			Sql = Sql & " TituloNoticia,"
			Sql = Sql & " Noticia"
				
			Sql = Sql & " ) "
			Sql = Sql & " Values "
			Sql = Sql & " ("
		
			Sql = Sql & "'" & xTituloNoticia & "',"
			Sql = Sql & "'" & xNoticia & "'"
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
</head>

<body background="img/Falogo03.gif">
<FORM action="RegistroNoticias.asp" method="post" name="frmReg" onSubmit="MM_validateForm('txt_TituloNoticia',,'Noticia');return document.MM_returnValue">
     
  <font color="#000080">
     
  <INPUT type="hidden" value="true" name="Comportamiento">
  </font><font color="#000080">
  <BR>
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
                    size=1><B><font color="#000080"><input type="text" name="txt_TituloNoticia" size="66"></font></B></FONT></TD>
        </TR>
      <TR>
        <TD height="51">
		<font face="Courier New">Noticia</font><font color="#000080"><BR>
		<TEXTAREA name=txt_Noticia cols=50 rows="13"></TEXTAREA></font></TD>
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
        <input type="submit" name="Submit" value="Enviar Noticia"></font></P></TD>
     
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
</font>
</body>
</html>