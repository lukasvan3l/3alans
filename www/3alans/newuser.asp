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

Welkom, <%=session("3alansname")%>!<br>
In dit scherm kun je de verschillende instellingen van <%=programName%> bewerken.<br>

<%select case request.querystring("a")
case "" 
	%>
	<form method="post" name="form" action="?a=cat_ink">
	<h2>Liquide middelen</h2>
	<table>
	<tr>
		<th></th>
		<th>Rekeningnummer</th>
		<th>Beginsaldo</th>
		<th>Datum</th>
	</tr>

	<%
	set rs = adocon.execute("SELECT * FROM liquide_middel WHERE userid="&session("3alansid")&" ORDER BY naam ASC")
	if not rs.EOF then
		set rs2= adocon.execute("SELECT * FROM overboeking WHERE userid="&session("3alansid")&" AND categorie='Beginsaldo' ORDER BY liquide_middel ASC")
		if rs("naam") = "Bank" and rs2("liquide_middel") = "Bank" then%>
			<tr>
				<td>Bankrekening</td>
				<td><input type="text" name="liq_bank" size=15 value="<%=rs("rekeningnummer")%>" disabled></td>
				<td>&euro; <input type="text" name="liq_bankstart" size=7 value="<%=formatnumber(rs2("bedrag"))%>" disabled></td>
				<td><input type="text" name="liq_bankdatum" size=10 value="<%=date2datum(rs2("datum"),false)%>" disabled></td>
			</tr>
			<%rs.movenext
			rs2.movenext
		end if
		if not rs.EOF then
			if rs("naam") = "Kas" and rs2("liquide_middel") = "Kas" then%>
				<tr>
					<td>Kas</td>
					<td></td>
					<td>&euro; <input type="text" name="liq_kasstart" size=7 value="<%=formatnumber(rs2("bedrag"))%>" disabled></td>
					<td><input type="text" name="liq_kasdatum" size=10 value="<%=date2datum(rs2("datum"),false)%>" disabled></td>
				</tr>
				<%rs.movenext
			rs2.movenext
			end if
		end if
		if not rs.EOF then
			if rs("naam") = "Spaarrekening" and rs2("liquide_middel") = "Spaarrekening" then%>
				<tr>
					<td>Spaarrekening</td>
					<td><input type="text" name="liq_spaar" size=15 value="<%=rs("rekeningnummer")%>" disabled></td>
					<td>&euro; <input type="text" name="liq_spaarstart" size=7 disabled></td>
					<td><input type="text" name="liq_spaardatum" size=10 value="<%=date2datum(now(),false)%>" disabled></td>
				</tr>
			<%end if
		end if
		rs.close
		rs2.close
	else
		%>
		<tr>
			<td>Bankrekening</td>
			<td><input type="text" name="liq_bank" size=15 value="<%=rs("rekeningnummer")%>"></td>
			<td>&euro; <input type="text" name="liq_bankstart" size=7></td>
			<td><input type="text" name="liq_bankdatum" size=10 value="<%=date2datum(now(),false)%>"></td>
		</tr>
		<tr>
			<td>Kas</td>
			<td></td>
			<td>&euro; <input type="text" name="liq_kasstart" size=7></td>
			<td><input type="text" name="liq_kasdatum" size=10 value="<%=date2datum(now(),false)%>"></td>
		</tr>
		<tr>
			<td>Spaarrekening</td>
			<td><input type="text" name="liq_spaar" size=15 value="<%=rs("rekeningnummer")%>"	></td>
			<td>&euro; <input type="text" name="liq_spaarstart" size=7></td>
			<td><input type="text" name="liq_spaardatum" size=10 value="<%=date2datum(now(),false)%>"></td>
		</tr>
		<%
	end if
	%>
	<tr>
		<td colspan=4 align="right" valign="bottom" height="30"><input type="submit" value="Volgende ->"></td>
	</tr>
	</table>
	<%set rs = nothing

'case "cat_ink"
	set rs = adocon.execute("SELECT * FROM liquide_middel WHERE userid="&session("3alansid"))
	if rs.EOF then
		adocon.execute("INSERT INTO liquide_middel(naam, rekeningnummer, userid) VALUES('Bank',"&insertToDb(request.form("liq_bank"),"float")&","&session("3alansid")&")")
		adocon.execute("INSERT INTO liquide_middel(naam, rekeningnummer, userid) VALUES('Kas',null,"&session("3alansid")&")")
		adocon.execute("INSERT INTO liquide_middel(naam, rekeningnummer, userid) VALUES('Spaarrekening','"&insertToDb(request.form("liq_spaar"),"float")&"',"&session("3alansid")&")")

		adocon.execute("INSERT INTO categorie(naam, meetellen, userid) VALUES('Beginsaldo',0,"&session("3alansid")&")")
		adocon.execute("INSERT INTO categorie(naam, meetellen, userid) VALUES('Kas',0,"&session("3alansid")&")")
		adocon.execute("INSERT INTO categorie(naam, meetellen, userid) VALUES('Bank',0,"&session("3alansid")&")")
		adocon.execute("INSERT INTO categorie(naam, meetellen, userid) VALUES('Spaarrekening',0,"&session("3alansid")&")")
		
		adocon.execute("INSERT INTO overboeking(datum, bedrag, aan_van, omschrijving, categorie, liquide_middel, userid) " & _
			"VALUES(FORMAT('" & request.form("liq_kasdatum") &"','dd-mm-yyyy')," & replace(request.form("liq_kasstart"),",",".") &",'Beginsaldo','Beginsaldo','Beginsaldo','Kas',"&session("3alansid")&")")
		adocon.execute("INSERT INTO overboeking(datum, bedrag, aan_van, omschrijving, categorie, liquide_middel, userid) " & _
			"VALUES(FORMAT('" & request.form("liq_bankdatum") &"','dd-mm-yyyy')," & replace(request.form("liq_bankstart"),",",".") &",'Beginsaldo','Beginsaldo','Beginsaldo','Bank',"&session("3alansid")&")")
		adocon.execute("INSERT INTO overboeking(datum, bedrag, aan_van, omschrijving, categorie, liquide_middel, userid) " & _
			"VALUES(FORMAT('" & request.form("liq_spaardatum") &"','dd-mm-yyyy')," & replace(request.form("liq_spaarstart"),",",".") &",'Beginsaldo','Beginsaldo','Beginsaldo','Spaarrekening',"&session("3alansid")&")")
	end if
	set rs = nothing

	if request.form("categorienform") = "true" then
		if not request.form("nieuwin") = "" then
			adocon.execute("INSERT INTO categorie(naam,meetellen,userid) VALUES("&insertToDb(request.form("nieuwin"),"string")&",1,"&session("3alansid")&")")
		end if
		if not request.form("nieuwuit") = "" then
			adocon.execute("INSERT INTO categorie(naam,meetellen,userid) VALUES("&insertToDb(request.form("nieuwuit"),"string")&",2,"&session("3alansid")&")")
		end if
		for each i in request.form()
			if (left(i,4) = "uit_" or left(i,4) = "ink_") and right(i,len(i)-4) <> request.form(i) then
				if request.form(i) = "" then
					set rs = adocon.execute("SELECT count(*) as aantal FROM overboeking WHERE categorie='" & right(i,len(i)-4) & "' and userid=" & session("3alansid"))
					if rs("aantal") = 0 then
						adocon.execute("DELETE FROM categorie WHERE naam=" & insertToDb(right(i,len(i)-4),"string")) & " AND userid=" & session("3alansid")
					else
						Response.write(i & " kan niet worden verwijderd (" & rs("aantal") & " verwijzingen).<br />")
					end if
				else
					adocon.execute("UPDATE overboeking SET categorie="&insertToDb(request.form(i),"string")&" WHERE categorie="&insertToDb(right(i,len(i)-4),"string")&" AND userid="&session("3alansid"))
					adocon.execute("UPDATE categorie SET naam="&insertToDb(request.form(i),"string")&" WHERE naam="&insertToDb(right(i,len(i)-4),"string")&" AND userid="&session("3alansid"))
				end if
			end if
		next
	end if
	%>
	<form method="post" name="form" action="?a=cat_ink">
	<input type=hidden value="true" name="categorienform" />
	<h2>Categorieen</h2>
	<table>
	<tr>
		<th>Inkomsten</th>
		<th>Uitgaven</th>
	</tr>
	<tr>
		<td valign=top>
			<% set rs = adocon.execute("SELECT * FROM categorie WHERE meetellen=1 AND userid="&session("3alansid")&" ORDER BY naam ASC")
			do until rs.EOF%>
				<input type="text" size=50 maxlength=50 value="<%=rs("naam")%>" name="ink_<%=rs("naam")%>" /><br />
			<%rs.movenext : loop
			rs.close : set rs = nothing%>
			<input type="text" size=50 maxlength=50 name="nieuwin" />
		</td>
		<td valign=top>
			<% set rs = adocon.execute("SELECT * FROM categorie WHERE meetellen=2 AND userid="&session("3alansid")&" ORDER BY naam ASC")
			do until rs.EOF%>
				<input type="text" size=50 maxlength=50 value="<%=rs("naam")%>" name="uit_<%=rs("naam")%>" /><br />
			<%rs.movenext : loop
			rs.close : set rs = nothing%>
			<input type="text" size=50 maxlength=50 name="nieuwuit" />
		</td>
	</tr>
	<tr>
		<td colspan=2 align="right" valign="bottom" height="30">
			<input type="submit" value="Opslaan">
		</td>
	</tr>
	<tr>
		<td colspan=2 align="right" valign="bottom" height="30">
			<input type="button" value="Volgende ->" onclick="window.location.href='newuser.asp?a=end'">
		</td>
	</tr>
	</table>

<%case "end"
	%>
	<form>
	<h2>
		Uw instellingen zijn gewijzigd!
	</h2>
	<%

end select%>

</form>
</body>
</html>
<% set adocon = nothing %>