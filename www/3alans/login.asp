<%	@LANGUAGE = VBScript %>
<%
Option Explicit
Response.Buffer = TRUE
Response.Expires = 0
session.lcid = 1043
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<!-- #include virtual="/3alans/asp/db.asp" -->
<!-- #include virtual="/3alans/asp/adovbs.inc" -->
<!-- #include virtual="/3alans/asp/functions.asp" -->

<%
if request.querystring("a") = "logout" then
	session("3alansname") = ""
	session("3alansid") = ""
	session.abandon
	response.redirect("default.asp")

elseif request.querystring("a") = "login" then
	set rs = adocon.execute("SELECT * FROM [user] WHERE verwijderd=false AND [username]='" & changequotes(request.form("username")) & "'")
	if rs.EOF then
		response.redirect("login.asp?a=f&goto=" & request.querystring("goto"))
	else
		if request.form("password") <> rs("password") then
			response.redirect("login.asp?a=f&goto=" & request.querystring("goto"))
		end if
	end if
	
	Response.Cookies("3alans") = request.form("username")
	Response.Cookies("3alans").Expires = Date + 365
	
	session("3alansname") = rs("username")
	session("3alansid") = rs("id")
	session.Timeout=20
	
	rs.close
	set rs = nothing
	
	set rs = execquery("SELECT * FROM overboeking WHERE userid="&session("3alansid"), adocon)
	if rs.EOF then
		response.redirect("newuser.asp")
	else
		response.redirect("default.asp")
	end if
end if
%>

<html>
<head>
	<title> <%=programName%> - Login</title>
	<meta name="author" content="Lukas van 3L" />
	<meta http-equiv="pragma" content="no-cache">
	<link rel="author" href="mailto:lukas@3l.nl" />
	<link rel="STYLESHEET" type="text/css" href="css/global.css">
</head>

<body>
<p>&nbsp;</p>
<H1 align="center"><%=programName%> - Login</H1>
<center>&nbsp;
<%if request.querystring("a") = "f" then%>
	Naam en wachtwoord komen niet overeen.
<%end if%>
</center>
<form action="login.asp?a=login&goto=<%= request.querystring("goto") %>" method="post" name="loginform">
<table width="200" align="center">
	<tr>
		<td>Naam:</td>
		<td><input type="text" name="username" value="<%=request.Cookies("3alans")%>" size="20" maxlength="20"></td>
	</tr>
	<tr>
		<td>Wachtwoord:</td>
		<td><input type="password" name="password" value="" size="20" maxlength="20"></td>
	</tr>
	<tr>
		<td></td>
		<td><input type="submit" value="log in"></td>
	</tr>
</table>
</form>

<%if request.Cookies("3alans") = "" then%>
	<!-- dit script moet onder het formulier zelf geplaats worden -->
	<script type="text/javascript" language="javascript">
		document.forms['loginform'].elements['username'].focus();
	</script>
<%else%>
	<!-- dit script moet onder het formulier zelf geplaats worden -->
	<script type="text/javascript" language="javascript">
		document.forms['loginform'].elements['password'].focus();
	</script>
<%end if%>

</body></html>
<% set adocon = nothing %>