<%

    if request.form("consulta") = "true" then
        openstr = "DSN=NOracl;UID=system;PWD=rdmcds;DBQ=ORCL ;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;"
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr

        sql = " SELECT * " & _
              " FROM Usuarios" & _
              " WHERE Usuario= '" & request.form("Usuario") & _
	      "' AND Clave ='" & request.form("Clave") & "'" 
			  
			  
           '			  response.Write "SQL--------->" & sql & "<---------" 
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, cn, 3, 3 

        if rs.eof  then
		    response.Redirect("mensaje0002.asp")
        else 
                        
                        Session("Usuario")= request.form("Usuario") 
                      
			     response.redirect("vertodos.asp") 
	end if
        
		rs.Close
		Set rs = Nothing
		
		cn.Close
		Set cn = Nothing

     END IF
%>
<HTML>
<HEAD>
<TITLE>Galenos Modulo de Administración</TITLE>
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

    <% if request.querystring("autorizado") = "falso" then
          response.redirect("mensaje0000100.asp")
       End if
     %> 

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

<FORM METHOD="Post" name="Login" ACTION="Login.asp">
    
  <div align="center"> 
    <input type="hidden" name="consulta" value="true">
    <font color="#FF2000" size="8" face="Lucida Sans"></font></div>
    
  <div align="center">
	  
    <TABLE height="91" BORDER=0 width="199">
      <TR>
        <TD class="style4 style5" width="4" rowspan="2">
		<img border="0" src="img/login.GIF" width="63" height="67"><TD class="style4" width="113"><INPUT NAME="Usuario" SIZE="15">
	  <TR>
        <TD class="style4" width="113"><INPUT TYPE="Password" NAME="Clave" SIZE="15">
	  <TR><TD COLSPAN=2 class="style4"><input name="Submit" type="Submit" value="Enviar">	    
    </TABLE>
  </div>
</FORM>
</BODY>
</HTML>
