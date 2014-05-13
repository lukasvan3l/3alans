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
<!-- #include virtual="/3alans/asp/header.asp" -->

<%if request.querystring("k") = "" then%>
	<h2>Overzichten</h2>
	<ul>
	<li><a href="?k=sd">Saldo per dag</li>
	<li><a href="?k=sm">Saldo per maand</li>
	<li><a href="?k=sq">Eigen query</li>
	<li><a href="?k=salvglma">Vandaag vergeleken met vorige maanden</li>
	<li><a href="?k=Grootboekrekeningen">Grootboekrekeningen (excel)</li>
	<li><a href="export.asp?k=maandoverzicht" target="_blank">Maandoverzichten (excel)</li>
	</ul>

<% elseif request.querystring("k") = "sq" then%>
	<h2>Saldoinformatie</h2>
	<%
	dim eenheid, hoogte, breedte, aantal, min, max, stap
	if (request.form("eenheid") <> "") then
		eenheid = request.form("eenheid")
		hoogte = request.form("hoogte")
		breedte = request.form("breedte")
		aantal = request.form("aantal")
		min = request.form("min")
		max = request.form("max")
		stap = request.form("stap")
		call saldoInf(eenheid, hoogte, breedte, aantal, min, max, stap)
	else
		eenheid = "d"
		hoogte = 10
		breedte = 980
		aantal = 365
		min = 4500
		max = 6000
		stap = 2
	end if
	%>
	<h2>Formulier</h2>
	<form method="post" action="?k=sq">
	<table>
	<tr>
		<td>Eenheid</td>
		<td><select name="eenheid">
			<option value="d" selected>Dag</option>
			<option value="m">Maand</option>
			</select>
		</td>
	</tr>
	<tr>
		<td>Hoogte</td>
		<td><input type="text" name="hoogte" value="<%=hoogte%>" /></td>
	</tr>
	<tr>
		<td>Breedte</td>
		<td><input type="text" name="breedte" value="<%=breedte%>" /></td>
	</tr>
	<tr>
		<td>Aantal</td>
		<td><input type="text" name="aantal" value="<%=aantal%>" /></td>
	</tr>
	<tr>
		<td>Minimum</td>
		<td><input type="text" name="min" value="<%=min%>" /></td>
	</tr>
	<tr>
		<td>Maximum</td>
		<td><input type="text" name="max" value="<%=max%>" /></td>
	</tr>
	<tr>
		<td>Stap</td>
		<td><input type="text" name="Stap" value="<%=stap%>" /></td>
	</tr>
	<tr>
		<td />
		<td><input type="submit" value="refresh" /></td>
	</tr>
	</form>
	<%
	
elseif request.querystring("k") = "sm" then%>
	<h2>Saldoinformatie per maand</h2>
	<%call saldoInf("m",8,900,36,4000,3000,1)

elseif request.querystring("k") = "sd" then%>
	<h2>Saldoinformatie per dag</h2>
	<%call saldoInf("d",8,900,180,4500,3000,1)

elseif request.querystring("k") = "salvglma" then
	dim ctDay
	if request.querystring("d") = "" then
		ctDay = day(now)
	else
		ctDay = request.querystring("d")
	end if
	for i = 1 to 31
		response.write("<a href='?k=salvglma&d="&i&"'>"&i&"</a> ")
	next
	%>
	<h2>Inkomsten en uitgaven op de <%=ctDay%>e</h2>
	<table>
	<tr><td>
	<table width=150>
	Inkomsten
	<%
	for i = -12 to 0
		call inUitPerDag(i,1,ctDay)
	next
	%>
	</table>
	</td><td width="10"></td><td>
	Uitgaven
	<table width=150>
	<%
	for i = -12 to 0
		call inUitPerDag(i,2,ctDay)
	next
	response.write("</table></td></tr></table>")

elseif request.querystring("k") = "Grootboekrekeningen" then
	%>
	<form method="post" action="export.asp" target="_blank" />
	<input type="hidden" name="type" value="grootboekrekeningen" />
	Vanaf datum: <input type="text" name="vanafdatum" value="" size="10" maxlength="10" />
	<br />
	<input type="submit" value="genereer excelbestand" />
	</form>

<%end if


sub saldoInf(eenheid, hoogte, breedte, aantal, min, max, stap)%>
	<table width=<%=breedte%> cellpadding=0 cellspacing=0 style="border:0px 1px solid #7C8FA1">
	<tr height="<%=max/hoogte%>">
	<%
	set rs = adocon.execute("SELECT sum(bedrag) AS startsaldo FROM overboeking WHERE categorie='Beginsaldo' and userid="&session("3alansid"))
	dim startsaldo, maxsaldo, minsaldo, mediumsaldo
	maxsaldo = -10000
	minsaldo = 10000
	mediumsaldo = 0
	teller = 0
	startsaldo = rs("startsaldo")
	rs.close : set rs = nothing
	for i = -1*aantal to 0 step stap
		datum = date2datum(dateadd(eenheid,i,now()),false)
		saldo = saldoOpDag(datum)
		%>
		<td valign="bottom" align="center">
			<a href="overboekingen.asp?m=<%=datum%>">
			<img 
				src="img/blue.gif" 
				width=<%=stap*0.7*breedte/aantal%>
				height="<%=(saldo-min)/hoogte%>"
				title="&euro; <%=saldo%> op <%=datum%>"></a>
		</td>
		<%
		if saldo <> startsaldo then
			teller = teller+1
			mediumsaldo = mediumsaldo + saldo
		end if
		if saldo > maxsaldo then
			maxsaldo = saldo
		end if
		if saldo < minsaldo then
			minsaldo = saldo
		end if
	next
	if teller = 0 then
		mediumsaldo = maxsaldo
	else
		mediumsaldo = mediumsaldo/teller
	end if
	%></tr>
	</table>
	
	<h2>Stats</h2>
	<table width=200>
	<tr>
		<td>Maximum</td>
		<td align="right">&euro; <%=formatnumber(maxsaldo)%></td>
	</tr>
	<tr>
		<td>Gemiddelde</td>
		<td align="right">&euro; <%=formatnumber(mediumsaldo)%></td>
	</tr>
	<tr>
		<td>Minimum</td>
		<td align="right">&euro; <%=formatnumber(minsaldo)%></td>
	</tr>
	<tr>
		<td>Huidig</td>
		<td align="right">&euro; <%=formatnumber(saldo)%></td>
	</tr>
	</table><%
end sub

sub inUitPerDag(dagi,inuit,dagnr)
	datum = date2datum(dateadd("m",dagi,dagnr&"-"&month(now)&"-"&year(now)),false)
	set rs = execquery("SELECT sum(bedrag) as saldo FROM overboeking LEFT JOIN categorie ON categorie.naam = overboeking.categorie WHERE overboeking.userid="&session("3alansid")&" AND categorie.userid="&session("3alansid")&" AND meetellen="&inuit&" AND (month(datum)="&month(datum)&" AND year(datum)=" & year(datum) & " AND day(datum)<=" & dagnr & ")",adocon)

	response.write("<tr><td><a href=""overboekingen.asp?m="&datum&""">")
	response.write(tweegetallen(month(datum)) & " "& year(datum)&"</a></td><td align='right'>&euro; ")
	if isnull(rs("saldo")) then
		response.write("0,00")
	else
		if inuit=2 then
			response.write(formatnumber(-1 * rs("saldo")))
		else
			response.write(formatnumber(rs("saldo")))
		end if
	end if
	response.write("</td></tr>")
	rs.close : set rs = nothing
end sub

set adocon = nothing%>
</body>
</html>