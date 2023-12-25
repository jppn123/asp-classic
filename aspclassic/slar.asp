<!DOCTYPE html>
<html>
<head>
	<meta charset="utf-8">
	<meta name="viewport" content="width=device-width, initial-scale=1">
	<title></title>
</head>
<body>
	<%
		if request.form <> "" then
			pnome = request.form("pnome")
			snome = request.form("snome")
			response.write "ola " & pnome & " " & snome & "<br>"
	%>
	<script type="text/javascript">
		console.log("achei");
	</script>
	<%
		end if

		dim fs, fo, x, foName
		set fs=Server.CreateObject("Scripting.FileSystemObject")
		set fo=fs.GetFolder("C:\inetpub\wwwroot\aspclassic\banners\")

		for each x in fo.files
			response.write "<img src= banners/" & x.name & " width=500 height=300>"
			response.write ""
		next
	%>

</body>
</html>

