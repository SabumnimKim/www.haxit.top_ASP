<%
Response.expires=0
Response.buffer=true

site = Request("site")

// ----- DATUM OCH TID ---------------------------------------------
dag = cstr(DatePart("d",Date()))
	If len(dag)< 2 Then
		dag = "0" +dag
	End if
manad = cstr(DatePart("m",Date()))
	If len(manad)<2 Then
		manad = "0" +manad
	End if
vecka = cstr(DatePart("w",Date()+2))
ar = cstr(DatePart("yyyy",Date()))
datum = dag&"."&manad&"."&ar
rdatum = ar&manad&dag
tid = FormatDateTime(time(),vbshorttime)

// ----- WEBBESÖKARE ------------------------------------------------
ipnummer = Request.ServerVariables("REMOTE_ADDR")
hostname = Request.ServerVariables("REMOTE_HOST")
lastpage = Request.ServerVariables("HTTP_REFERER")
If inStr(lastpage,"?")>0 Then
	lastpage = left(lastpage,inStr(lastpage,"?")-1)
End if


%>