<% 'GEMENSAMMA FUNKTIONER 			'(för samtliga sidor på webbplatsen) %>
<!--#include file="_f_getname.asp" -->


<%
'CSS-DEKLARATIONER				'Alla CSS-filer som skall finnas. Separera med ett ","-tecken
css = 	pathcss & "index.css,"  &_
	pathcss & "top.css,"    &_
	pathcss & "left.css,"   &_
	pathcss & "main.css,"	&_
	pathcss & "right.css,"	&_
	pathcss & "bottom.css," &_
	""

'JAVASCRIPT-DEKLARATIONER			'Alla JAVASCRIPT-filer som skall finnas. Separera med ett ","-tecken
javascript=	""

%>