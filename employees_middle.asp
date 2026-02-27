<%
dblocation="c:\InetPub\haxit.top\mdb\"				'Fysisk adress till databasen 	
photolocation="c:\InetPub\haxit.top\employee_photos\"		'Fysisk adress till anställdas foton
letter = Request("letter")

callpage ="employees.asp"
referer = Request.ServerVariables("HTTP_REFERER")
referer = right(referer,len(callpage))
If NOT referer=callpage THEN
	Response.write "Error 4 - Referer not allowed to call this page"
	RESPONSE.END 
END IF



If Not letter ="" Then

' SKAPA DATABAS-KONTAKT OCH ÖPPNA DATABASEN
Set ConDatabas		= Server.CreateObject("ADODB.Connection")
ConDatabas.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ="&dblocation&"Employees.mdb"
Set RS      	     = Server.CreateObject("ADODB.Recordset")
RS.CursorType        = 0
RS.LockType          = 1

' LÄS ALLA EMPLOYEES IN I ARRAYEN
Sql="SELECT * FROM Employees ORDER BY Nickname ASC"
Set RS = ConDatabas.Execute(Sql)
employeesString = ""
While Not RS.Bof And Not RS.Eof	
	employeesString = employeesString &","& RS("Nickname")
	RS.MoveNext
Wend

' SKAPA ARRAY FÖR ALLA EMPLOYEES
employeesArray = split(employeesString,",")

%>
<div class="employeeletter"> <% =letter %> </div><br>
<%
' LOOPA GENOM ARRAYEN OCH SKRIV UT EMPLOYEES SOM MATCHAR BOKSTAVEN
for each i in employeesArray

display = "no"
If letter="All" Then
	If Not i = "" Then
	Sql="SELECT * FROM Employees Where Nickname='"&i&"'"
	Set RS = ConDatabas.Execute(Sql)
	display = "yes"
	End if
Else
	If letter = Left(i,1) Then
		Sql="SELECT * FROM Employees Where Nickname='"&i&"'"
		Set RS = ConDatabas.Execute(Sql)
		display = "yes"
	End If
End if

If display="yes" Then
%>
<table border="0" cellspacing="0" cellpadding="0">
 <tr>
  <td rowspan="6" width="65" align="left" valign="top"><% If RS("ServiceStatus")="Honorable Discharge" Then Response.Write("<img src='images/medal_honorable_discharge.png'>") End If %></td>
  <td rowspan="6" width="110" align="left" valign="top"><img src="employee_photos/

<%
If RS("Photo") = "" Then
	Response.write("default.jpg")
Else
	' GRANSKA OM BILDEN FINNS
	dim fs
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	if fs.FileExists(photolocation & RS("Photo")) then
		Response.write RS("Photo")
	Else
		Response.write("default.jpg")
	End If
	set fs=nothing
End If
%>

" width="100"></td>
  <td class="employeetdname" align="left" valign="bottom"><a href="javascript:openEmployeeData('<% =RS("ID") %>')"><% =RS("Nickname") %></a></td>
  <td rowspan="6" width="65" align="right" valign="top" align="center">
	<% 
	If Not RS("Medals")="" Then
		For j = 1 to RS("Medals")
		Response.write("<img src='images/medal" & j & ".gif' style='margin:2px;'><br>")			
		Next
	End If 
	%>
  </td>
 </tr>
 <tr>
  <td class="employeetdinfo" align="left" valign="top">Nickname</td>
 </tr>
 <tr>
  <td class="employeetdstatus" align="left" valign="bottom" bgcolor="

<%If RS("ServiceStatus") = "Active" Then Response.Write("#99ccff") End If %>
<%If RS("ServiceStatus") = "Missing" Then Response.Write("#ffff99") End If %>
<%If RS("ServiceStatus") = "Killed in Action" Then Response.Write("#ffcccc") End If %>
<%If RS("ServiceStatus") = "Honorable Discharge" Then Response.Write("#99cc99") End If %>
<%If RS("ServiceStatus") = "Passive" Then Response.Write("#cccccc") End If %>


"><% =RS("ServiceStatus") %></td>
 </tr>
 <tr>
  <td class="employeetdinfo" align="left" valign="top">Service Status</td>
 </tr>
 <tr>
  <td class="employeetdname" align="left" valign="bottom"><% =RS("SecurityAccessLevel") %></td>
 </tr>
 <tr>
  <td class="employeetdinfo" align="left" valign="top">Security Access Level</td>
 </tr>
 <tr>
  <td colspan="4" height="30"></td>
 </tr>
</table>
<%


display = "no"
End if
Next

' STÄNG DATABAS-KONTAKTEN
Set RS = nothing
ConDatabas.Close
Set ConDatabas = nothing

End if
%>


