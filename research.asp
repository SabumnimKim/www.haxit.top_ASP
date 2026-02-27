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

'EGNA CSS-DEKLARATIONER			'Separera med ","-tecken för varje fil
css = css&","&	pathcss & "research.css,"  &_		
		""

'EGNA JAVASCRIPT-DEKLARATIONER		'Alla JAVASCRIPT-filer som skall finnas. Separera med ett ","-tecken
javascript = javascript&","&	pathscripts & "Research_openObjectID.js," &_
				pathscripts & "Research_tooltip.js," &_
				""
%>

<!--#include file="header.asp" -->

<!--#include file="top.asp" -->

<%
' SKAPA DATABAS-KONTAKT OCH ÖPPNA DATABASEN
Set ConDatabas		= Server.CreateObject("ADODB.Connection")
ConDatabas.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ="&dblocation&"Research.mdb"
Set RS      	     = Server.CreateObject("ADODB.Recordset")
RS.CursorType        = 0
RS.LockType          = 1
%>

<div class="researchframe">

<img src onerror="javascript:openObjectID('')">

	<div class="researchleft">
		<a href="index.asp"><img src="<% =pathimages %>logo.png" width="250" border="0"></a>
		<br>
		<div class="researchheading">RESEARCH OBJECTS:</div>
		<br>
		<table border="0" cellspacing="1" cellpadding="0">


<%
' SKRIV UT ALLA RESEARCH OBJECTS
Sql="SELECT * FROM ResearchObjects ORDER BY number"
Set RS = ConDatabas.Execute(Sql)
While Not RS.Bof And Not RS.Eof	

	'RÄKNA ANTALET ENTRIES FÖR RESPEKTIVE RESEARCH OBJECT
	researchObject = RS("number")
	Sql2="SELECT COUNT(*) AS antal FROM ResearchEntries WHERE researchObject = '"&researchObject&"'"
	Set RS2 = ConDatabas.Execute(Sql2)
	antal = RS2("antal")
%>
		 <tr>
		  <td class="researchidtd"><a href="javascript:openObjectID('<% =RS("number") %>')"><% =Replace(RS("number"),"-","#") %></a></td>
		  <td class="researchobjecttd"><a href="javascript:openObjectID('<% =RS("number") %>')"><% =RS("name") %></a> <span class="researchcount">( <% =antal %>)</span></td>
		 </tr>

<%
antal = 0
RS.MoveNext
Wend
%>
		 </table>
	</div>
	<div id="researchmiddle" class="researchmiddle">
	</div>

<%

' STÄNG DATABAS-KONTAKTEN
Set RS2 = nothing
Set RS = nothing
ConDatabas.Close
Set ConDatabas = nothing
%>

<!--#include file="bottom.asp" -->

<!--#include file="footer.asp" -->
