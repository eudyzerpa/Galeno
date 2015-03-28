<%
 
       xemail = Session("xemail")
 
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
		 
				
	    
		
		 sqlupdate = " UPDATE registro " & _
                    " SET Especialidad = '" & Request.Form("Txt_Especialidad") & "'," & _ 
                      		    " SubEspecialidad = '" & Request.Form("Txt_SubEspecialidad")  & "'," & _
                       		    " HospitalTrabaja = '" & Request.Form("Txt_HospitalTrabaja")  & "'," & _
				    			" Cargo = '" & Request.Form("Txt_Cargo")  & "'," & _
				    			" ClinicaTrabaja = '" & Request.Form("Txt_ClinicaTrabaja")  & "'," & _
				    			" Consultorio = '" & Request.Form("Txt_Consultorio")  & "'," & _
				      			" Telefono = '" & Request.Form("Txt_Telefono")  & "'," & _
							    " Fax = '" & Request.Form("Txt_Fax")  & "'," & _
				                " Celular = '" & Request.Form("Txt_Celular")  & "'," & _
                                " FechaCumpleanos = '" & Now & "'," & _
								" TemasInteres = '" & Request.Form("Txt_TemasInteres") & "'," & _
                       		    " RecibirInformacion = '" & Request.Form("Txt_RecibirInformacion") & "' " & _ 
                      		    " WHERE email = '" & xemail & "'"
         
         cn.Execute sqlupdate, raffected
         
         if raffected > 0 then
              response.Redirect("mensaje0003.asp")
         else
           response.Redirect("mensaje000034.asp")
         end if
     
      end if
                     
			
        
		rs.Close
		Set rs = Nothing
		
		cn.Close
		Set cn = Nothing
		
		

%>