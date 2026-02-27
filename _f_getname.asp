<%
Function getname(inparam)
	dblocation="c:\InetPub\haxit.top\mdb\"			'Fysisk adress till databasen 
	employeeCode = inparam
	name = ""

	' SKAPA DATABAS-KONTAKT OCH ÖPPNA DATABASEN
	Set ConDatabasX		= Server.CreateObject("ADODB.Connection")
	ConDatabasX.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ="&dblocation&"Employees.mdb"
	Set RSX      	     = Server.CreateObject("ADODB.Recordset")
	RSX.CursorType        = 0
	RSX.LockType          = 1

	SqlX="SELECT * FROM Employees where employeeCode ='"&employeeCode&"'"
	Set RSX = ConDatabasX.Execute(SqlX)

	If Not RSX.Eof Then name = RSX("nickName") End If

	' STÄNG DATABAS-KONTAKTEN
	Set RSX = nothing
	ConDatabasX.Close
	Set ConDatabasX = nothing

	getname = name
End Function
%>