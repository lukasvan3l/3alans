<%	@LANGUAGE = VBScript %>
<%
Option Explicit
Response.Buffer = TRUE
Response.Expires = 0
session.lcid = 1043

response.ContentType = "application/vnd.ms-excel"

%>

<!-- #include virtual="/3alans/asp/db.asp" -->
<!-- #include virtual="/3alans/asp/adovbs.inc" -->
<!-- #include virtual="/3alans/asp/functions.asp" -->

<html>
<head>
	<title> <%=programName%> </title>
	<meta name="author" content="Lukas van 3L" />
	<meta http-equiv="pragma" content="no-cache" />
	<link rel="author" href="mailto:lukas@3l.nl" />
	<link href="css/global.css" rel="stylesheet" type="text/css" />
	<script language="javascript" src="js/functions.js" type="text/javascript"></script>
</head>

<body>

<%if request.form("type") = "grootboekrekeningen" then%>
	<h2>Kasboeken (vanaf <%=request.form("vanafdatum")%>)</h2>
	<%
	set rs = adocon.execute("SELECT * FROM categorie WHERE userid="&session("3alansid")&" ORDER BY naam ASC")
	do until rs.EOF
		%>
		<h3><%=rs("naam")%></h3>
		<table width="800" cellspacing=0 cellpadding=2 style="border:0px 1px solid #7C8FA1">
		<tr class="header">
			<td width="100">Datum</td>
			<td width="75" align="right" style="padding-right: 10px;">Bedrag</td>
			<td width="200">Aan / Van</td>
			<td width="200">Omschrijving</td>
			<td width="75">Middel</td>
			<td width="75">Afschrift</td>
			<td width="75">Bon</td>
		</tr>
		<%
		dim bedragliqmid, query
		totaal = 0
		if (request.form("vanafdatum") <> "") then 
			query = "SELECT * FROM overboeking WHERE userid="&session("3alansid")&" AND categorie='"&rs("naam")&"' AND datum > "+inserttodb(request.form("vanafdatum"),"date")+" ORDER BY datum ASC"
		else
			query = "SELECT * FROM overboeking WHERE userid="&session("3alansid")&" AND categorie='"&rs("naam")&"' ORDER BY datum ASC"
		end if
		set rs2 = adocon.execute(query)
		do until rs2.EOF
			%>
			<tr>
				<td nowrap valign="top">
					&nbsp;<%=date2datum(rs2("datum"),false)%>
				</td>
				<td align="right" valign="top" style="padding-right: 10px;" nowrap>
					<%=replace(formatnumber(rs2("bedrag")),",",".")%>
					<%totaal = totaal + rs2("bedrag")%>
				</td>
				<td valign="top">
					<%=rs2("aan_van")%>
				</td>
				<td valign="top">
					<%=left(rs2("omschrijving"),40)%>
				</td>
				<td valign="top">
					<%=rs2("liquide_middel")%>
				</td>
				<td valign="top">
					<%=rs2("afrekeningnr")%>
				</td>
				<td nowrap valign="top">
					<%=rs2("bonnr")%>
				</td>
			</tr>
			<%
		rs2.movenext : loop
		rs2.close : set rs2 = nothing
		%>
		
		<tr>
			<td nowrap>Totaal:</td>
			<td align="right" style="padding-right: 10px;" nowrap>
				<%=totaal%>
			</td>
			<td colspan=5></td>
		</tr>

		</table>
		<%

	rs.movenext : loop
	rs.close : set rs = nothing



elseif request.querystring("k") = "maandoverzicht" then
	%><table cellspacing=0 cellpadding=2 style="border:0px 1px solid #7C8FA1">
	<tr><%
	dim maandenTonen, aantalmaanden
	aantalmaanden = 11

	if request.querystring("mt") = "" then
		maandenTonen = -10
	else
		maandenTonen = request.querystring("mt")
	end if

	set rs = execquery("SELECT * FROM categorie WHERE NOT meetellen=0 AND userid="&session("3alansid")&" ORDER BY naam asc", adocon)
	for i = maandenTonen to maandenTonen+aantalmaanden-1
		maand = date2datum(dateadd("m",i,now()),false)%>
		<td valign="top">
		<%
		totaaluit = 0
		totaalin = 0
		euriuit = ""
		euriin = ""
		do until rs.EOF
			set rs2 = execquery("SELECT SUM(bedrag) as saldo FROM overboeking WHERE categorie='"&rs("naam")&"' AND month(datum)="&month(maand)& " AND userid="&session("3alansid"), adocon)
			if rs("meetellen") = 2 then
				euriuit = euriuit & "<tr onmouseover=""hlhr(this,'over')"" onmouseout=""hlhr(this,'out')"" onclick=""window.location.href='overboekingen.asp?c="&rs("naam")&"&m="&maand&"'"" style='cursor:pointer;'><td>"&rs("naam")&"</td><td align='right'>"
				if not isnull(rs2("saldo")) then
					totaaluit = totaaluit - rs2("saldo")
					euriuit = euriuit & formatnumber(-1 * rs2("saldo"))
				else
					euriuit = euriuit & "0,00"
				end if
				euriuit = euriuit & "</td></tr>"
			elseif rs("meetellen") = 1 then
				euriin = euriin & "<tr onmouseover=""hlhr(this,'over')"" onmouseout=""hlhr(this,'out')"" onclick=""window.location.href='overboekingen.asp?c="&rs("naam")&"&m="&maand&"'"" style='cursor:pointer;'><td>"&rs("naam")&"</td><td align='right'>"
				if not isnull(rs2("saldo")) then
					totaalin = totaalin + rs2("saldo")
					euriin = euriin & formatnumber(rs2("saldo"))
				else
					euriin = euriin & "0,00"
				end if
				euriin = euriin & "</td></tr>"
			end if
			set rs2 = nothing
		rs.movenext : loop
		rs.movefirst%>
		<p>Inkomsten <a href="overboekingen.asp?m=<%=month(maand) & " " & year(maand)%>"><%maandvanjaar(month(maand))%></a></p>
		<table width="200" cellspacing=0 cellpadding=2>
		<%=euriin%>
		<tr>
			<td></td>
			<td align="right" style="border-top:1px solid #7C8FA1; height: 25px" valign="top">
				<strong>&euro; <%=formatnumber(totaalin)%></strong>
			</td>
		</tr>
		</table>
		<P>Uitgaven <a href="overboekingen.asp?m=<%=month(maand) & " " & year(maand)%>"><%maandvanjaar(month(maand))%></a></P>
		<table width="200" cellspacing=0 cellpadding=2>
		<%=euriuit%>
		<tr>
			<td></td>
			<td align="right" style="border-top:1px solid #7C8FA1; height: 25px" valign="top">
				<strong>&euro; <%=formatnumber(totaaluit)%></strong>
			</td>
		</tr>
		</table>
		</td>
	<%next
	rs.close : set rs = nothing
	%></table><%

end if
%>

</body></html>