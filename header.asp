<html>
<head>
<title>HA.ax/IT - haxit</title>

<!--#include file="header_metadata.asp" -->

<%
' CSS-UTSKRIFT (från variabeln "css")
If Not css="" Then

	' SPLITTAR VARIABELN TILL EN ARRAY (enligt ","-tecknet)
	cssArray = split(css,",")

	' LOOPAR GENOM ARRAYEN OCH SKRIVER UT EN HTML-RAD PER INPUT I ARRAYEN
	for ijk=0 to UBound(cssArray)
		if NOT cssArray(ijk)="" Then
			Response.write "<link rel=""stylesheet"" href=""" & cssArray(ijk) & """>" & vbCrLf
		end if
	next
End if

' JAVASCRIPT-UTSKRIFT (från variablen "javascript")
If Not javascript="" Then

	' SPLITTAR VARIABELN TILL EN ARRAY (enligt ","-tecknet)
	javascriptArray = split(javascript,",")

	' LOOPAR GENOM ARRAYEN OCH SKRIVER UT EN HTML-RAD PER INPUT I ARRAYEN
	for ijk=0 to UBound(javascriptArray)
		if NOT javascriptArray(ijk)="" Then
			Response.write "<script language=""JavaScript"" src="""& JavaScriptArray(ijk) & """></script>" & vbCrLf
		end if
	next
End if
%>

</head>
<body>

<% '-------- ONMOUSEOVER INFO ------ %>


<% '-------- STARTAR BOTTENLAGRET ------ %>
<div style="position:absolute;z-index:1;width:100%;" align="center">