function openRegistryLetter(letter)
{
	const xhttp = new XMLHttpRequest();
	xhttp.onload = function() 
	{
		document.getElementById("employeesmiddle").innerHTML = this.responseText;
	}
	xhttp.open("GET", "employees_middle.asp?letter="+letter+"", true);
	xhttp.send();
}