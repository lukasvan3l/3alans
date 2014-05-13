<%	@LANGUAGE = VBScript %>
<%
Option Explicit
Response.Buffer = TRUE
Response.Expires = 0
session.lcid = 1043

dim subtotaal
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<!-- #include virtual="/3alans/asp/db.asp" -->
<!-- #include virtual="/3alans/asp/adovbs.inc" -->
<!-- #include virtual="/3alans/asp/functions.asp" -->
<!-- #include virtual="/3alans/asp/header.asp" -->

<h2>Overzicht huidige situatie</h2>

<table width="300" cellspacing=0 cellpadding=2>
<tr onmouseover="hlhr(this,'over')" onmouseout="hlhr(this,'out')" onclick="window.location.href='overboekingen.asp?l=kas'" style="cursor:pointer;">
	<td>Saldo kas</td>
	<td align="right">&euro; 
		<%
		set rs = execquery("SELECT SUM(bedrag) as saldo FROM overboeking WHERE liquide_middel='Kas' AND userid="&session("3alansid"), adocon)
		set rs2 = execquery("SELECT SUM(bedrag) as saldo FROM overboeking WHERE categorie='Kas' AND userid="&session("3alansid"), adocon)
			if isnull(rs2("saldo")) then	
				subtotaal = rs("saldo")
			else
				subtotaal = rs("saldo")-rs2("saldo")
			end if
			Response.write(formatnumber(subtotaal))
			totaal = totaal + subtotaal
		set rs = nothing
		set rs2 = nothing
		%>
	</td>
</tr>
<tr onmouseover="hlhr(this,'over')" onmouseout="hlhr(this,'out')" onclick="window.location.href='overboekingen.asp?l=bank'" style="cursor:pointer;">
	<td>Saldo bank</td>
	<td align="right">&euro; 
		<%
		set rs = execquery("SELECT SUM(bedrag) as saldo FROM overboeking WHERE liquide_middel='Bank' AND userid="&session("3alansid"), adocon)
			if isnull(rs("saldo")) then	
				subtotaal = 0
			else
				subtotaal = rs("saldo")
			end if
			Response.write(formatnumber(subtotaal))
			totaal = totaal + subtotaal
		set rs = nothing
		%>
	</td>
</tr>
<tr onmouseover="hlhr(this,'over')" onmouseout="hlhr(this,'out')" onclick="window.location.href='overboekingen.asp?l=spaarrekening'" style="cursor:pointer;">
	<td>Saldo spaarrekening</td>
	<td align="right">&euro; 
		<%
		set rs = execquery("SELECT SUM(bedrag) as saldo FROM overboeking WHERE liquide_middel='Spaarrekening' AND userid="&session("3alansid"), adocon)
		set rs2 = execquery("SELECT SUM(bedrag) as saldo FROM overboeking WHERE categorie='Spaarrekening' AND userid="&session("3alansid"), adocon)
			if isnull(rs2("saldo")) then	
				subtotaal = rs("saldo")
			else
				subtotaal = rs("saldo")-rs2("saldo")
			end if
			if isnull(subtotaal) then
				subtotaal = 0
			end if
			Response.write(formatnumber(subtotaal))
			totaal = totaal + subtotaal
		set rs = nothing
		set rs2 = nothing
		%>
	</td>
</tr>
<tr>
	<td></td>
	<td align="right" style="border-top:1px solid #7C8FA1; height: 25px">
		<strong>&euro; <%=formatnumber(totaal)%></strong>
	</td>
</tr>
</table>

<h2>Specificatie per categorie</h2>

<table width="900" cellspacing=0 cellpadding=2 style="border:0px 1px solid #7C8FA1">
<tr>
<%
dim maandenTonen, aantalmaanden
aantalmaanden = 4

if request.querystring("mt") = "" then
	maandenTonen = -3
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
		set rs2 = execquery("SELECT SUM(bedrag) as saldo FROM overboeking WHERE categorie='"&rs("naam")&"' AND month(datum)="&month(maand)&" AND year(datum)="&year(maand)&" AND userid="&session("3alansid"), adocon)
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
rs.close : set rs = nothing%>
</tr>
<tr>
	<td colspan=4 align="center">
	<a href="default.asp?mt=<%=maandenTonen-1%>">&lt;&lt Geschiedenis</a> ::
	<a href="default.asp?mt=<%=maandenTonen+1%>">Toekomst &gt;&gt;</a>
	</td>
</tr>
</table>

</body>
</html>
<%set adocon = nothing%>