<%
dblocation="c:\InetPub\haxit.top\mdb\"				'Fysisk adress till databasen 	
photolocation="c:\InetPub\haxit.top\employee_photos\"		'Fysisk adress till anställdas foton
id = Request("id")

callpage ="employees.asp"
referer = Request.ServerVariables("HTTP_REFERER")
referer = right(referer,len(callpage))
If NOT referer=callpage THEN
	Response.write "Error 4 - Referer not allowed to call this page"
	RESPONSE.END 
END IF


If Not id ="" Then

' SKAPA DATABAS-KONTAKT OCH ÖPPNA DATABASEN
Set ConDatabas		= Server.CreateObject("ADODB.Connection")
ConDatabas.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ="&dblocation&"Employees.mdb"
Set RS      	     = Server.CreateObject("ADODB.Recordset")
RS.CursorType        = 0
RS.LockType          = 1

Set RS2      	     = Server.CreateObject("ADODB.Recordset")
RS2.CursorType        = 0
RS2.LockType          = 1

'LÄS DATABASVÄRDEN FÖR VALD EMPLOYEE
Sql="SELECT * FROM Employees Where ID="&id&""
Set RS = ConDatabas.Execute(Sql)

Sql2="SELECT * FROM PersonnelRegistry Where employeeCode='"&RS("employeeCode")&"'"
Set RS2 = ConDatabas.Execute(Sql2)

If RS("SecurityAccessLevel") = "C" Then 
	securityImage="security_card_level_C.jpg"
	bgcolorclass="employeesecuritybgC"
End if
If RS("SecurityAccessLevel") = "B" Then 
	securityImage="security_card_level_B.jpg" 
	bgcolorclass="employeesecuritybgB"
End if
If RS("SecurityAccessLevel") = "A" Then 
	securityImage="security_card_level_A.jpg" 
	bgcolorclass="employeesecuritybgA"
End if
If RS("SecurityAccessLevel") = "A1" Then 
	securityImage="security_card_level_A1.jpg" 
	bgcolorclass="employeesecuritybgA1"
End if
If RS("SecurityAccessLevel") = "A2" Then 
	securityImage="security_card_level_A2.jpg" 
	bgcolorclass="employeesecuritybgA2"
End if
If RS("SecurityAccessLevel") = "A3" Then 
	securityImage="security_card_level_A3.jpg" 
	bgcolorclass="employeesecuritybgA3"
End if
If RS("SecurityAccessLevel") = "A4" Then 
	securityImage="security_card_level_A4.jpg" 
	bgcolorclass="employeesecuritybgA4"
End if

%>

<div class="employeefile">Security Access Card:</div>
<br>

<table border="0" cellspacing="0" cellpadding="0">
 <tr>
  <td width="300" height="450" valign="top">
	<div class="employeeinfocard"><img src="images/<% =securityImage %>" width="300"></div>
	<div class="employeeinfotext">
		<table border="0" cellspacing="0" cellpadding="0">
		 <tr>
		  <td colspan="3" width="300" height="113"></td>
		 </tr>
		 <tr>
		  <td width="131" height="170" valign="top">
			<table border="0" cellspacing="0" cellpadding="0">
			 <tr>
			  <td rowspan="4" width="10" height="170"></td>
			  <td width="106" height="120" valign="top"><img src="employee_photos/<% =RS("photo")%>" width="106" height="120" style="border:1px solid black;border-collapse: collapse;"></td>
			  <td rowspan="4" width="15" height="170"></td>
			  </td>
			 </tr>
			 <tr>
			  <td width="105" height="15"></td>
			 </tr>
			 <tr>
			  <td width="105" height="20" align="center" valign=middle" class="<% =bgcolorclass %>"><% =RS("employeeCode") %></td>
			 </tr>
			 <tr>
			  <td width="105" height="15"></td>
			 </tr>
			</table>
		  </td>
		  <td height="170" width="147" valign="top">
			<table border="0" cellspacing="0" cellpadding="0">
			 <tr>
			  <td colspan="3" width="147" height="22" valign="bottom" class="<% =bgcolorclass %>"><% =RS("Nickname") %></td>
			 </tr>
			 <tr>
			  <td colspan="3" width="147" height="10"></td>
			 </tr>
			 <tr>
			  <td colspan="3" width="147" height="19" valign="bottom" class="<% =bgcolorclass %>"><% If Not RS2.Eof Then Response.write(RS2("dateOfBirth")) End If %></td>
			 </tr>
			 <tr>
			  <td colspan="3" width="147" height="10"></td>
			 </tr>
			 <tr>
			  <td width="70" height="20" valign="bottom" class="<% =bgcolorclass %>"><% If Not RS2.Eof Then Response.write(RS2("sex")) End If %></td>
			  <td width="5"></td>
			  <td width="72" height="20" valign="bottom" class="<% =bgcolorclass %>"><% If Not RS2.Eof Then Response.write(RS2("bloodType")) End If %></td>
			 </tr>
			 <tr>
			  <td colspan="3" width="147" height="10"></td>
			 </tr>
			 <tr>
			  <td width="70" height="20" valign="bottom" class="<% =bgcolorclass %>"><% If Not RS2.Eof Then Response.write(RS2("height")) End If %></td>
			  <td width="5"></td>
			  <td width="72" height="20" valign="bottom" class="<% =bgcolorclass %>"><% If Not RS2.Eof Then Response.write(RS2("weight")) End If %></td>
			 </tr>
			 <tr>
			  <td colspan="3" width="147" height="10"></td>
			 </tr>
			 <tr>
			  <td colspan="3" width="147" height="19" valign="bottom" class="<% =bgcolorclass %>"><% If Not RS2.Eof Then Response.write(RS2("department")) End If %></td>
			 </tr>
			 <tr>
			  <td colspan="3" width="147" height="10"></td>
			 </tr>
			 <tr>
			  <td colspan="3" width="147" height="19" valign="bottom" class="<% =bgcolorclass %>"><% If Not RS2.Eof Then Response.write(RS2("rank")) End If %></td>
			 </tr>
			</table>
		  </td>
		  <td width="23"></td>
		 </tr>
		</table>
	</div>
  </td>
 </tr>
</table>

<div class="employeefile">Personnel File:</div>
<br>

<div style="background-image: url('../images/personnelfile_bg01.png');">
	<table border="0" cellspacing="0" cellpadding="0">
	 <tr>
	  <td width="65" height="149" rowspan="2"></td>
	  <td width="156" height="99"></td>
	  <td width="115" height="149" rowspan="2"></td>
	  <td width="164" height="149" rowspan="2" valign="top" align="center">
		<div style="position:relative;">
		<img src="employee_photos/<% =RS("photo")%>" height="140" style="border:1px solid black; top:5px; position:relative; z-index:4;">
		<img src="images/paperclip.png" style="position:absolute; top:0px; left:20px; z-index:5;">
		</div>
	  </td>
	 </tr>
	 <tr>
	  <td width="156" height="50" style="font-size: 10px;" align="left" valign="top">
		<b>Name: </b><% If Not RS2.Eof Then Response.write(RS2("name")) End If %><br>
		<b>Employee Code: </b><% If Not RS2.Eof Then Response.write(RS2("employeeCode")) End If %><br>
		<b>Signature Date: </b><% If Not RS2.Eof Then Response.write(RS2("signatureDate")) End If %><br>
	  </td>
	 </tr>
	</table>
</div>
<table border="0" cellspacing="0" cellpadding="0">
 <tr>
  <td width="500" height="20"><img src="images/personnelfile_bg02.png"></td>
 </tr>
</table>
<div style="background-image: url('../images/personnelfile_bg03.png');">
	<table border="0" cellspacing="0" cellpadding="0">
	 <tr>
	  <td width="65"></td>
	  <td width="385" style="font-size: 10px;" align="left" valign="top">
		<b>Background: </b><br><p></p><% If Not RS2.Eof Then Response.write(Replace(RS2("background"),vbcrlf,"<br><p></p>")) End If %><br>
	  </td>
	  <td width="50" ></td>
	 </tr>
	</table>
</div>
<table border="0" cellspacing="0" cellpadding="0">
 <tr>
  <td width="500" height="29"><img src="images/personnelfile_bg04.png"></td>
 </tr>
</table>
<div style="background-image: url('../images/personnelfile_bg05.png');">
	<table border="0" cellspacing="0" cellpadding="0">
	 <tr>
	  <td width="65"></td>
	  <td width="385" style="font-size: 10px;" align="left" valign="top">
		<b>Strengths: </b><br><p></p><% If Not RS2.Eof Then Response.write(Replace(RS2("strengths"),vbcrlf,"<br><p></p>")) End If %><br>
	  </td>
	  <td width="50" ></td>
	 </tr>
	</table>
</div>
<table border="0" cellspacing="0" cellpadding="0">
 <tr>
  <td width="500" height="39"><img src="images/personnelfile_bg06.png"></td>
 </tr>
</table>
<div style="background-image: url('../images/personnelfile_bg07.png');">
	<table border="0" cellspacing="0" cellpadding="0">
	 <tr>
	  <td width="65"></td>
	  <td width="385" style="font-size: 10px;" align="left" valign="top">
		<b>Weaknesses: </b><br><p></p><% If Not RS2.Eof Then Response.write(Replace(RS2("weaknesses"),vbcrlf,"<br><p></p>")) End If %><br>
	  </td>
	  <td width="50" ></td>
	 </tr>
	</table>
</div>
<div style="background-image: url('../images/personnelfile_bg08.png');">
	<table border="0" cellspacing="0" cellpadding="0">
	 <tr>
	  <td width="65" height="95"></td>
	  <td width="385" style="font-size: 10px;" align="left" valign="top">
		<br><p></p>
		<b>Date in: </b><% If Not RS.Eof Then Response.write(RS("date_in")) End If %><br>
		<b>Date out: </b><% If Not RS.Eof Then Response.write(RS("date_out")) End If %><br>
	  </td>
	  <td width="50" ></td>
	 </tr>
	</table>
</div>
<%
' STÄNG DATABAS-KONTAKTEN
Set RS2 = nothing
Set RS = nothing
ConDatabas.Close
Set ConDatabas = nothing

End if
%>


