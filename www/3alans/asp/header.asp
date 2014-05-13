<html>
<head>
	<title><%=programName%></title>
	<meta name="author" content="Lukas van 3L" />
	<meta http-equiv="pragma" content="no-cache" />
	<link rel="author" href="mailto:lukas@3l.nl" />
	<link href="css/global.css" rel="stylesheet" type="text/css" />
	<script language="javascript" src="js/functions.js" type="text/javascript"></script>
	<link href="/favicon.ico" rel="shortcut icon" />
</head>

<body>

<div style="position:absolute; right:0px; top:10px;">
	<a href="default.asp">Home</a> :: 
	<a href="saldo.asp">Overzichten</a> :: 
	<a href="import.asp">Toevoegen</a> ::
	<a href="overboekingen.asp?bw=true<%
	for each i in request.querystring
		if i <> "bw" then
			response.write("&" & i & "=" & request.querystring(i))
		end if
	next
	%>">Bewerken</a> ::
	<a href="newuser.asp">Instellingen</a> ::
	<a href="login.asp?a=logout">Log uit</a>
</div>

<h1> <%=programName%> </h1>