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
css = css&","&	pathcss & "index.css,"  &_		
		""

'EGNA JAVASCRIPT-DEKLARATIONER		'Alla JAVASCRIPT-filer som skall finnas. Separera med ett ","-tecken
'javascript = javascript&","&	pathscripts & "search.js," &_
'				""
%>

<!--#include file="header.asp" -->

<!--#include file="top.asp" -->

<% 



%>

<table border="0" cellspacing="0" cellpadding="0">
 <tr>
  <td width="800" height="50"></td>
 </tr>
 <tr>
  <td width="800" height="121" align="center"><img src="images/ha_logo.gif" alt="IT-programmet vid Högskolan på Åland"></td>
 </tr>
 <tr>
  <td width="800" height="15"></td>
 </tr>
 <tr>
  <td width="800" height="349" align="center"><img src="images/haxit_frontlogo.gif" alt="?haxit?"></td>
 </tr>
 <tr>
  <td width="800" height="15"></td>
 </tr>
</table>

<table border="0" cellspacing="1" cellpadding="0">
 <tr>
  <td class="indextdbutton"><a href="employees.asp">EMPLOYEE REGISTRY</a></td>
  <td class="indextdbutton"><a href="research.asp">RESEARCH</a></td>
  <td class="indextdbutton"><a href="researchfacilities.asp">RESEARCH FACILITIES</a></td>
 </tr>
</table>

<table border="0" cellspacing="1" cellpadding="0">
 <tr>
  <td width="800" height="40"></td>
 </tr>
</table>

<table border="0" cellspacing="1" cellpadding="0">
 <tr>
  <td align="center" valign="middle"><img src="images/mini_re5.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_umbrella.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_tricell.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_starwars.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_simpsons.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_fedora.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_pfsense.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_wolfenstein.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_cs.gif" style="margin:5px;"></td>
 </tr>
</table>
<table border="0" cellspacing="1" cellpadding="0">
 <tr>
  <td align="center" valign="middle"><img src="images/mini_transformers.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_discord.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_marvel.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_c64.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_wy.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_raspberrypi.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_vscode.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_php.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_adobe.gif" style="margin:5px;"></td>
 </tr>
</table>
<table border="0" cellspacing="1" cellpadding="0">
 <tr>
  <td align="center" valign="middle"><img src="images/mini_github.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_mysql.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_intel.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_microsoft.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_android.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_wireshark.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_southpark.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_quake.gif" style="margin:5px;"></td>
  <td align="center" valign="middle"><img src="images/mini_kotipizza.gif" style="margin:5px;"></td>
 </tr>
</table>

<!--#include file="bottom.asp" -->

<!--#include file="footer.asp" -->
