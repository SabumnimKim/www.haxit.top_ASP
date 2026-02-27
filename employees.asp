<% checksession  = "0"			'0=nej, 1=ja %>
<% checklastpage = "0"			'0=nej, 1=ja %>
<% securitylevel = "0" 			'0=ej inloggad, 1=user, 2=poweruser, 3=admin %>
<!--#include file="_security.asp" -->

<!--#include file="_config.asp" -->
<!--#include file="_globals.asp" -->


<% 'GEMENSAMMA FUNKTIONER (för samtliga sidor på webbplatsen) %>
<!--#include file="_public.asp" -->


<% 'EGNA FUNKTIONER (endast för denna sida) %>

<%
'VARIABELDEKLARATIONER
pagename = "index.asp"			'Sidans namn eller närmaste klick-sida
title = ""				'Sidans titel på svenska

'EGNA CSS-DEKLARATIONER			'Serarera med ","-tecken för varje fil
css = css&","&	pathcss & "employees.css,"  &_		
		""

'EGNA JAVASCRIPT-DEKLARATIONER		'Alla JAVASCRIPT-filer som skall finnas. Separera med ett ","-tecken
javascript = javascript&","&	pathscripts & "Employees_openRegistryLetter.js," &_
				pathscripts & "Employees_openEmployeeData.js," &_

				""
%>

<!--#include file="header.asp" -->

<!--#include file="top.asp" -->

<%
' SKAPA DATABAS-KONTAKT OCH ÖPPNA DATABASEN
Set ConDatabas		= Server.CreateObject("ADODB.Connection")
ConDatabas.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ="&dblocation&"Employees.mdb"
Set RS      	     = Server.CreateObject("ADODB.Recordset")
RS.CursorType        = 0
RS.LockType          = 1

' KOLLA HUR MÅNGA EMPLOYEES DET FINNS I DATABASEN
'Sql="SELECT COUNT(Nickname) AS antal FROM Employees"
'Set RS = ConDatabas.Execute(Sql)
'antal = RS("antal")

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

Function employeeLetter(letter)
	count = 0
	for each i in employeesArray
		If letter="all" Then 
			count = count + 1
		Else
			If letter = Left(i,1) Then
				count = count + 1
			End If
		End If
	next
     Response.write count
End Function

' STÄNG DATABAS-KONTAKTEN
Set RS = nothing
ConDatabas.Close
Set ConDatabas = nothing
%>

<div class="employeesframe">

	<div class="employeesleft">
		<a href="index.asp"><img src="<% =pathimages %>logo.png" width="250" border="0"></a>
		<br>
		<div class="employeeheading">EMPLOYEE REGISTRY:</div>
		<br>

		<table border="0" cellspacing="1" cellpadding="0">
		 <tr>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('A')">A</a> <span class="employeecount">(<% =employeeLetter("A") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('B')">B</a> <span class="employeecount">(<% =employeeLetter("B") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('C')">C</a> <span class="employeecount">(<% =employeeLetter("C") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('D')">D</a> <span class="employeecount">(<% =employeeLetter("D") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('E')">E</a> <span class="employeecount">(<% =employeeLetter("E") %>)</span></td>
		 </tr>
		 <tr>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('F')">F</a> <span class="employeecount">(<% =employeeLetter("F") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('G')">G</a> <span class="employeecount">(<% =employeeLetter("G") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('H')">H</a> <span class="employeecount">(<% =employeeLetter("H") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('I')">I</a> <span class="employeecount">(<% =employeeLetter("I") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('J')">J</a> <span class="employeecount">(<% =employeeLetter("J") %>)</span></td>
		 </tr>
		 <tr>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('K')">K</a> <span class="employeecount">(<% =employeeLetter("K") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('L')">L</a> <span class="employeecount">(<% =employeeLetter("L") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('M')">M</a> <span class="employeecount">(<% =employeeLetter("M") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('N')">N</a> <span class="employeecount">(<% =employeeLetter("N") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('O')">O</a> <span class="employeecount">(<% =employeeLetter("O") %>)</span></td>
		 </tr>
		 <tr>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('P')">P</a> <span class="employeecount">(<% =employeeLetter("P") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('Q')">Q</a> <span class="employeecount">(<% =employeeLetter("Q") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('R')">R</a> <span class="employeecount">(<% =employeeLetter("R") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('S')">S</a> <span class="employeecount">(<% =employeeLetter("S") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('T')">T</a> <span class="employeecount">(<% =employeeLetter("T") %>)</span></td>
		 </tr>
		 <tr>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('U')">U</a> <span class="employeecount">(<% =employeeLetter("U") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('V')">V</a> <span class="employeecount">(<% =employeeLetter("V") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('X')">X</a> <span class="employeecount">(<% =employeeLetter("X") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('Y')">Y</a> <span class="employeecount">(<% =employeeLetter("Y") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('Z')">Z</a> <span class="employeecount">(<% =employeeLetter("Z") %>)</span></td>
		 </tr>
		 <tr>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('Å')">Å</a> <span class="employeecount">(<% =employeeLetter("Å") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('Ä')">Ä</a> <span class="employeecount">(<% =employeeLetter("Ä") %>)</span></td>
		  <td class="employeetd"><a href="javascript:openRegistryLetter('Ö')">Ö</a> <span class="employeecount">(<% =employeeLetter("Ö") %>)</span></td>
		  <td class="employeetd"></td>
		  <td class="employeetd"></td>
		 </tr>
		 <tr>
		  <td colspan="5" height="10"></td>
		 </tr>
		 <tr>
		  <td colspan="5" align="center">
			<table border="0" cellspacing="1" cellpadding="0">
			 <tr>
			  <td class="employeetdall"><a href="javascript:openRegistryLetter('All')">Show all</a> <span class="employeecount">(<% =employeeLetter("all") %>)</span></td>
			 </tr>
			</table>
		  </td>
		 </tr>
		</table>
		<br>
	</div>

	<div id="employeesmiddle" class="employeesmiddle">
	</div>

	<div id="employeesright" class="employeesright">
	</div>


</div>

<!--#include file="bottom.asp" -->

<!--#include file="footer.asp" -->
