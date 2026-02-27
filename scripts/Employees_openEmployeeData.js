function openEmployeeData(id)
{

	const xhttp = new XMLHttpRequest();
	xhttp.onload = function() 
	{
		document.getElementById("employeesright").innerHTML = this.responseText;
	}
	xhttp.open("GET", "employees_right.asp?id="+id+"", true);
	xhttp.send();
}