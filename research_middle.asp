<!--#include file="_f_getname.asp" -->

<%
dblocation="c:\InetPub\haxit.top\mdb\"			'Fysisk adress till databasen 	
id = Request("id")

callpage ="research.asp"
referer = Request.ServerVariables("HTTP_REFERER")
referer = right(referer,len(callpage))
If NOT referer=callpage THEN
	Response.write "Error 4 - Referer not allowed to call this page"
	RESPONSE.END 
END IF

' SKAPA DATABAS-KONTAKT OCH ÖPPNA DATABASEN
Set ConDatabas		= Server.CreateObject("ADODB.Connection")
ConDatabas.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ="&dblocation&"Research.mdb"
Set RS      	     = Server.CreateObject("ADODB.Recordset")
RS.CursorType        = 0
RS.LockType          = 1
Set RS2      	     = Server.CreateObject("ADODB.Recordset")
RS2.CursorType        = 0
RS2.LockType          = 1
Set RS3      	     = Server.CreateObject("ADODB.Recordset")
RS3.CursorType        = 0
RS3.LockType          = 1
%>

<div class="researchheading" style="height:50px;">VIRUS STRAIN MAP:</div><br>
<div style="position:relative;">

	<div style="z-index:2;">
	<img src="images/virusmap.png" usemap="#virusmap">
	<map name="virusmap">
	  <area shape="circle" coords="655,58,23" alt="Stairway of the Sun" href="javascript:openObjectID('TCL-00')"  onMouseOver="showTooltip(event, 'TCL-00')" onMouseOut="hideTooltip(event, 'TCL-00')">
	  <area shape="circle" coords="283,116,23" alt="Uroboros" href="javascript:openObjectID('TCL-04')" onMouseOver="showTooltip(event, 'TCL-04')" onMouseOut="hideTooltip(event, 'TCL-04')">
	  <area shape="circle" coords="662,127,23" alt="Progenitor Virus" href="javascript:openObjectID('TCL-01')" onMouseOver="showTooltip(event, 'TCL-01')" onMouseOut="hideTooltip(event, 'TCL-01')">
	  <area shape="circle" coords="1023,119,23" alt="T-Veronica Virus" href="javascript:openObjectID('TCL-03')" onMouseOver="showTooltip(event, 'TCL-03')" onMouseOut="hideTooltip(event, 'TCL-03')">
	  <area shape="circle" coords="511,240,23" alt="Tyrant Virus" href="javascript:openObjectID('TCL-02')" onMouseOver="showTooltip(event, 'TCL-02')" onMouseOut="hideTooltip(event, 'TCL-02')">
	  <area shape="circle" coords="167,346,23" alt="T-JCCC203" href="javascript:openObjectID('TCL-06')" onMouseOver="showTooltip(event, 'TCL-06')" onMouseOut="hideTooltip(event, 'TCL-06')">
	  <area shape="circle" coords="760,313,23" alt="G-Virus" href="javascript:openObjectID('TCL-05')" onMouseOver="showTooltip(event, 'TCL-05')" onMouseOut="hideTooltip(event, 'TCL-05')">
	  <area shape="circle" coords="981,315,23" alt="C-Virus" href="javascript:openObjectID('TCL-07')" onMouseOver="showTooltip(event, 'TCL-07')" onMouseOut="hideTooltip(event, 'TCL-07')">
	  <area shape="circle" coords="163,628,23" alt="Las Plagas" href="javascript:openObjectID('TCL-10')" onMouseOver="showTooltip(event, 'TCL-10')" onMouseOut="hideTooltip(event, 'TCL-10')">
	  <area shape="circle" coords="228,611,23" alt="T-Abyss Virus" href="javascript:openObjectID('TCL-09')" onMouseOver="showTooltip(event, 'TCL-09')" onMouseOut="hideTooltip(event, 'TCL-09')">
	  <area shape="circle" coords="662,673,23" alt="A-Virus" href="javascript:openObjectID('TCL-11')" onMouseOver="showTooltip(event, 'TCL-11')" onMouseOut="hideTooltip(event, 'TCL-11')">
	</map>
	</div>

<%
' INGET RESEARCH OBJECT HAR VALTS
If id ="" Then
	' SKAPA TOOLTIP-DIV FÖR ALLA RESEARCH OBJECTS
	Sql="SELECT * FROM ResearchObjects ORDER BY number"
	Set RS = ConDatabas.Execute(Sql)
	While Not RS.Bof And Not RS.Eof	

		'RÄKNA ANTALET ENTRIES FÖR RESPEKTIVE RESEARCH OBJECT
		researchObject = RS("number")
		Sql2="SELECT COUNT(*) AS antal FROM ResearchEntries WHERE researchObject = '"&researchObject&"'"
		Set RS2 = ConDatabas.Execute(Sql2)
		antal = RS2("antal")
	%>
		<div id="<% =RS("number") %>" class="tooltip">
		<img src="images/biohazard.png" width="100"style="padding-bottom:5px;"><br>
		<b>Number:</b> <% =Replace(RS("number"),"-","#") %> <br>   vbCrLF
		<b>Name:</b> <% =RS("name") %> <br>
		<b>Research entries:</b> <% =antal %><br>
		<img src="research_photos/<% =RS("picture") %>" width="200" style="padding-top:5px; padding-bottom:5px;">
		<br>
		<% =RS("description") %>
		</div>
	<%
	antal = 0
	RS.MoveNext
	Wend

Else
' ETT RESEARCH OBJECT HAR VALTS

	' LÄS ALLA VARIABLER FÖR AKTUELLT RESEARCH OBJECT
	Sql="SELECT * FROM ResearchObjects WHERE number='"&id&"'"
	Set RS = ConDatabas.Execute(Sql)

	creator = RS("creator")
	number = RS("number")
	description = Replace(RS("description"),vbcrlf,"<br><p></p>")
	reference = RS("reference")

	'RÄKNA ANTALET ENTRIES FÖR RESPEKTIVE RESEARCH OBJECT
	researchObject = RS("number")
	Sql2="SELECT COUNT(*) AS antal FROM ResearchEntries WHERE researchObject = '"&researchObject&"'"
	Set RS2 = ConDatabas.Execute(Sql2)
	antal = RS2("antal")
	%>

	<div style="background-color: rgba(0, 0, 0, 0.7);color:#ffffff;width:1141;height:1000px;position:absolute; padding-top:50px;top:0px; left:0px; z-index:5;">
	<table border="0" cellspacing="0" cellpadding="0">
	 <tr>
	  <td width="846" height="114" style="background-image: url('images/research_bg01.png');">
		<table border="0" cellspacing="0" cellpadding="0">
		 <tr>
		  <td colspan="4" height="20"></td>
		 </tr>
		 <tr>
		  <td width="60" valign="top" align="center" class="researchclose">|<a href="javascript:openObjectID('')" style="">X</a>|</td>
		  <td width="310">
			<table border="0" cellspacing="0" cellpadding="0">
			 <tr>
			   <td width="310" height="40" valign="middle" style="font-family:Arial;font-size:20px;font-weight:bold;"> <% =Replace(id,"-","#") %></td>
			 </tr>
			 <tr>
			  <td width="310" height="30" valign="top" style="font-family:Arial;font-size:20px;font-style:italic;"> <% =RS("name") %></td>
			 </tr>
			</table>
		  </td>
		  <td width="100"></td>
		  <td width="300" valign="top" align="left">
			<table border="0" cellspacing="0" cellpadding="0">
			 <tr>
			  <td width="300" height="22" valign="middle" style="font-family:Arial;font-size:12px;"><b>File created: </b><% =RS("date") %> | <% =RS("time") %></td>
			 </tr>
			 <tr>
			   <td width="300" height="22" valign="middle" style="font-family:Arial;font-size:12px;"><b>By: </b><% =getname(creator) %> (<% =creator %>) </td>
			 </tr>
			 <tr>
			   <td width="300" height="22" valign="middle" style="font-family:Arial;font-size:12px;"><b>Research Entries: </b><% =antal %></td>
			 </tr>
			</table>
		  </td>
		 </tr>
		</table>
	  </td>
	 </tr>
	 <tr>
	  <td width="846" style="background-image: url('images/research_bg03.png');">
		<table border="0" cellspacing="0" cellpadding="0">
		 <tr>
		  <td width="60"></td>
		  <td width="310">
			<table border="0" cellspacing="0" cellpadding="0">
			 <tr>
			  <td width="300" height="35" valign="top" style="padding-bottom:10px;"><img src="research_photos/<% =number %>.png" width="200" style="border:1px solid;border-color:#000000;"></td>
			 </tr>
			 <tr>
			  <td width="300" height="35" valign="top" style="padding-bottom:10px;" class="researchtext"><% =description %></td>
			 </tr>
			 <tr>
			  <td width="300" height="35" valign="top" style="padding-bottom:10px;" class="researchreference"><a href="<% =reference %>"><% =reference %></a></td>
			 </tr>
			</table>
		  </td>
		  <td width="100"></td>
		  <td width="300" valign="top">
			<table border="0" cellspacing="0" cellpadding="0">
			<%
			' LÄS ALLA RESEARCH ENTRIES FÖR AKTUELLT RESEARCH OBJECT

			youtubeimagep1 = "https://img.youtube.com/vi/"
			youtubeimagep2 = "/mqdefault.jpg"

			Sql3="SELECT * FROM ResearchEntries WHERE researchObject = '"&id&"'"
			Set RS3 = ConDatabas.Execute(Sql3)
			While Not RS3.Bof And Not RS3.Eof	

				link = RS3("link")
				videoid = Mid(link,InStrRev(link,"/")+1,len(link)-InStrRev(link,"/"))
				youtubeimage = youtubeimagep1&videoid&youtubeimagep2
				
				%>	
				<tr>
				<td width="150" align="left" valign="top" height="110"><a href="<% =link %>"><img src="<% =youtubeimage %>" width="140" height="100"></a></td>
				<td width="150" align="left" valign="top">
					<span class="researchentrydate"><% =RS3("date") %> | <% =RS3("time") %><br></span>
					<span class="researchentrytext"><% =RS3("text") %><br><p></p> </span>
					<span class="researchentryauthors">
					<% 
						arrayauthors = split(RS3("authors"),",")
						For Each author In arrayauthors
							Response.write getname(author) &"("& author &")<br>"
						Next
					 %></span>
				</td>
				</tr>
				<tr>
				<td colspan="2" height="22" valign="top"></td>
				</tr>			
			<%
			RS3.MoveNext
			Wend
			%>
			</table>
		  </td>
		 </tr>
		</table>
	  </td>
	 </tr>
	 <tr>
	  <td width="846" height="76" style="background-image: url('images/research_bg04.png');">
	 </tr>
	</table>
	</div>
</div>
<%


End if

' STÄNG DATABAS-KONTAKTEN
Set RS3 = nothing
Set RS2 = nothing
Set RS = nothing
ConDatabas.Close
Set ConDatabas = nothing
%>