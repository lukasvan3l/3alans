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

<%
if request.querystring("a") = "saveChanges" then
	dim arrDelete
	for each i in request.form
		if right(i,6) = "delete" and request.form(i) = "on" then
			arrDelete = arrDelete & i & ";"
		elseif right(i,5) = "datum" then
			arrImport = arrImport & i & ";"
		end if
	next

	arrImport = replace(arrImport,"_datum","")
	arrImport = split(arrImport,";")

	arrDelete = replace(arrDelete,"_delete","")
	arrDelete = split(arrDelete,";")

	for each i in arrDelete
		if not i = "" then
			strQuery = "DELETE FROM overboeking WHERE id=" & insertToDb(i,"int") & " AND userid=" & insertToDb(session("3alansid"),"int")

			response.write(strQuery & "<br>")
			adocon.execute(strQuery)
		end if
	next

	for each i in arrImport
		if not i = "" then
			strQuery = "UPDATE overboeking SET "
			strQuery = strQuery & "datum="&insertToDb(request.form(i&"_datum"),"date")&", "
			strQuery = strQuery & "bedrag=" & insertToDb(request.form(i&"_bedrag"),"float") & ", "
			strQuery = strQuery & "aan_van=" & insertToDb(request.form(i&"_aan_van"),"string") & ", "
			strQuery = strQuery & "omschrijving=" & insertToDb(request.form(i&"_omschrijving"),"string") & ", "
			strQuery = strQuery & "categorie=" & insertToDb(request.form(i&"_categorie"),"string") & ", "
			strQuery = strQuery & "bonnr=" & insertToDb(request.form(i&"_bonnr"),"int") & ", "
			strQuery = strQuery & "afrekeningnr=" & insertToDb(request.form(i&"_afrekeningnr"),"string") & ", "
			strQuery = strQuery & "liquide_middel=" & insertToDb(request.form(i&"_liquide_middel"),"string") & " "
			strQuery = strQuery & "WHERE id=" & insertToDb(i,"int") & " AND userid=" & insertToDb(session("3alansid"),"int")
			
			response.write(strQuery & "<br>")
			if not (localhost=true) then
				adocon.execute(strQuery)
			end if
		end if
	next
	%>

	<h2>Update geslaagd</h2>
	De wijzigingen zijn succesvol opgeslagen in de database.<br>

	<a href="overboekingen.asp<%
	for each i in request.querystring
		if i <> "bw" and i <> "a" then
			response.write("&" & i & "=" & request.querystring(i))
		end if
	next%>">Klik hier om terug te gaan naar het overzicht.</a>
	<%


else


	if not request.querystring("m") = "" then
		maand = month(request.querystring("m"))
		jaar = year(request.querystring("m"))
	end if
	categorie = request.querystring("c")
	liqmid = request.querystring("l")
	%>
	<table width=900 cellspacing=0 cellpadding=2>
	<tr>
		<td><h2>Overzicht overboekingen</h2></td>
		<td align="right">

			<select onchange="window.location.href='overboekingen.asp?c='+this.options[this.selectedIndex].value+'&m=<%=Request.querystring("m")%>&l=<%=request.querystring("l")%>'" id="ddlbCategorie">
			<option value="">- kies categorie -</option>
			<%
			set rs2 = execquery("select naam from categorie WHERE userid="&session("3alansid")&" order by naam asc", adocon)
			do until rs2.EOF
				if rs2("naam") = request.querystring("c") then
					response.write("<option selected value='"&rs2("naam")&"'>"&rs2("naam")&"</option>")
				else
					response.write("<option value='"&rs2("naam")&"'>"&rs2("naam")&"</option>")
				end if
			rs2.movenext : loop
			rs2.close : set rs2 = nothing
			%>
			</select>

			<select onchange="window.location.href='overboekingen.asp?l='+this.options[this.selectedIndex].value+'&m=<%=Request.querystring("m")%>&c=<%=request.querystring("c")%>'" id="ddlbLiqmid">
			<option value="">- kies liquide middel -</option>
			<%
			set rs2 = execquery("select naam from liquide_middel WHERE userid="&session("3alansid")&" order by naam asc", adocon)
			do until rs2.EOF
				if lcase(rs2("naam")) = lcase(request.querystring("l")) then
					response.write("<option selected value='"&rs2("naam")&"'>"&rs2("naam")&"</option>")
				else
					response.write("<option value='"&rs2("naam")&"'>"&rs2("naam")&"</option>")
				end if
			rs2.movenext : loop
			rs2.close : set rs2 = nothing
			%>
			</select>

			<select onchange="window.location.href='overboekingen.asp?m='+this.options[this.selectedIndex].value+'&c=<%=Request.querystring("c")%>&l=<%=request.querystring("l")%>'" id="ddlbDate">
			<option value="">- kies maand -</option>
			<%
			set rs2 = execquery("select min(datum) as mindate, max(datum) as maxdate FROM overboeking WHERE userid="&session("3alansid")&" order by 1 asc", adocon)
			datum = rs2("mindate")
			do while datum <= rs2("maxdate")
				dim selectedMonth
				if request.querystring("m") <> "" then
					selectedMonth = month(request.querystring("m")) & " " & year(request.querystring("m"))
				end if
				if month(datum) & " "& year(datum) = selectedMonth then
					response.write("<option selected value='" & month(datum) & " "& year(datum) &"'>"&month(datum) & " "& year(datum) &"</option>")
				else
					response.write("<option value='" & month(datum) & " "& year(datum) &"'>" & month(datum) & " "& year(datum) &"</option>")
				end if
				datum = dateadd("m",1,datum)
			loop
			rs2.close : set rs2 = nothing
			%>
			</select>

		</td>
	</tr>
	</table>

	<form method="POST" action="?a=saveChanges<%
	for each i in request.querystring
		if not(i = "a" or i = "bw") then
			response.write("&" & i & "=" & request.querystring(i))
		end if
	next%>" name="form">
	<table width="900" cellspacing=0 cellpadding=2 style="border:0px 1px solid #7C8FA1">
	<tr class="header">
		<td>Datum</td>
		<td>Bon</td>
		<td align="right" style="padding-right: 10px;">Bedrag</td>
		<td>Aan / Van</td>
		<td>Omschrijving</td>
		<td>Categorie</td>
		<td>Middel (afschr)</td>
	</tr>
	<%
	dim qryString
	qryString = "SELECT "
	if liqmid="" and categorie="" and maand="" then
		qryString = qryString & "TOP 20 "
	end if
	qryString = qryString & "* FROM overboeking WHERE isnull(bedrag) = false "
	if not liqmid = "" then
		qryString = qryString + "AND (liquide_middel='"&lcase(liqmid)&"' OR categorie='"&lcase(liqmid)&"') "
	end if
	if not categorie = "" then
		qryString = qryString + "AND (categorie='"&lcase(categorie)&"') "
	end if
	if not maand = "" then
		qryString = qryString + "AND (month(datum)="&maand&" AND year(datum)=" & jaar & ") "
	end if
	qryString = qryString + "AND (userid="&session("3alansid")& ") ORDER BY datum DESC"
	set rs = execquery(qryString, adocon)
	
	'response.write(qryString)
	
	if request.querystring("bw") = "true" then
		bewerk = true
	else
		bewerk = false
	end if
	do until rs.EOF
		dim bedragliqmid, bewerk
		if lcase(rs("categorie")) = lcase(liqmid) then
			bedragliqmid = -1*rs("bedrag")
		else
			bedragliqmid = rs("bedrag")
		end if
		%>
		<tr<%if (bedragliqmid > 0 and categorie = "") then 
			response.write("style=""background-color: #E5E5E5;""") 
		end if%>>
			<td nowrap>
				<%if bewerk=true then%>
					<input name="<%=rs("id")%>_delete" type="checkbox">
					<input name="<%=rs("id")%>_datum" type="text" maxlength=10 size="10" value="<%=date2datum(rs("datum"),false)%>">
				<%else%>
					<%=date2datum(rs("datum"),false)%>
				<%end if%>
			</td>
			<td>
				<%if bewerk=true then%>
					<input name="<%=rs("id")%>_bonnr" type="text" maxlength=3 size=3 value="<%=rs("bonnr")%>">
				<%else%>
					<%=rs("bonnr")%>
				<%end if%>
			</td>
			<td align="right" style="padding-right: 10px;">
				<%if bewerk=true then%>
					<input name="<%=rs("id")%>_bedrag" type="text" maxlength=12 size="10" value="<%=replace(formatnumber(rs("bedrag")),".","")%>" style="text-align:right;">
				<%else%>
					&euro; <%=formatnumber(bedragliqmid)%>
				<%end if
				totaal = totaal + bedragliqmid%>
			</td>
			<td>
				<%if bewerk=true then%>
					<input name="<%=rs("id")%>_aan_van" type="text" maxlength=50 size="20" value="<%=rs("aan_van")%>">
				<%else%>
					<%=rs("aan_van")%>
				<%end if%>
			</td>
			<td>
				<%if bewerk=true then%>
					<input name="<%=rs("id")%>_omschrijving" type="text" maxlength=50 size="40" value="<%=rs("omschrijving")%>">
				<%else%>
					<%=left(rs("omschrijving"),40)%>
				<%end if%>
			</td>
			<td>
				<%if bewerk=true then%>
					<select name="<%=rs("id")%>_categorie"><option>- geen -</option>
					<%
					set rs2 = execquery("select naam from categorie WHERE userid="&session("3alansid")&" order by naam asc", adocon)
					do until rs2.EOF
						if rs2("naam") = rs("categorie") then
							response.write("<option selected>"&rs2("naam")&"</option>")
						else
							response.write("<option>"&rs2("naam")&"</option>")
						end if
					rs2.movenext : loop
					rs2.close : set rs2 = nothing
					%>
					</select>
				<%else%>
					<a href="?a=boek&c=<%=rs("categorie")%>"><%=rs("categorie")%></a>
				<%end if%>
			</td>
			<td>
				<%if bewerk=true then%>
					<select name="<%=rs("id")%>_liquide_middel"><option>- geen -</option>
					<%
					set rs2 = execquery("select naam from liquide_middel WHERE userid="&session("3alansid")&" order by naam asc", adocon)
					do until rs2.EOF
						if rs2("naam") = rs("liquide_middel") then
							response.write("<option selected>"&rs2("naam")&"</option>")
						else
							response.write("<option>"&rs2("naam")&"</option>")
						end if
					rs2.movenext : loop
					rs2.close : set rs2 = nothing
					%>
					</select>
					<input type="text" name="<%=rs("id")%>_afrekeningnr" maxlength=8 size=8 value="<%=rs("afrekeningnr")%>" />
				<%else%>
					<a href="?a=boek&l=<%=rs("liquide_middel")%>"><%=rs("liquide_middel")%></a>
					<%if rs("afrekeningnr") <> "" then
						response.write("<em>(" & rs("afrekeningnr") & ")</em>")
					end if
				end if%>
			</td>
		</tr>
		<%
	rs.movenext: loop
	rs.close : set rs = nothing
	%>
	<tr>
		<td colspan=2 />
		<td align=right style="border-top:1px solid #7C8FA1; padding-right:10px;">&euro; <%=formatnumber(totaal)%></td>
		<td colspan=3 />
		<td><%if bewerk = true then response.write("<input type='submit' value='Opslaan'>") end if%></td>
	</tr>
	</table>
	</form>
<%end if

set adocon = nothing%>
</body></html>