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
if request.querystring("a") = "" then
	if localhost = false then
		Dim UploadProgress, PID, barref
		Set UploadProgress = Server.CreateObject("Persits.UploadProgress")
		PID = "PID=" & UploadProgress.CreateProgressID()
		barref = "framebar.asp?to=10&" & PID
		%>
		<SCRIPT LANGUAGE="JavaScript">
		function ShowProgress()
		{
			strAppVersion = navigator.appVersion;
			if (document.uploadForm.Path.value != "")
			{
				if (strAppVersion.indexOf('MSIE') != -1 && strAppVersion.substr(strAppVersion.indexOf('MSIE')+5,1) > 4)
				{
					winstyle = "dialogWidth=375px; dialogHeight:130px; center:yes";
					window.showModelessDialog('<% = barref %>&b=IE',null,winstyle);
				}
				else
				{
				window.open('<% = barref %>&b=NN','','width=370,height=115', true);
				}
			}
			return true;
		}
		</SCRIPT>
		<h2>Importeer gegevens</h2>
		<form action="?a=import_2&<% = PID %>" method="post" enctype="multipart/form-data" name="uploadForm" id="uploadForm" onsubmit="return ShowProgress();">
		<blockquote>
			<select name="filetype">
				<option>Rabobank</option>
				<option>Handmade</option>
			</select>
			<input type="file" name="Path" size="40"> <input type="submit" value="Upload">
		</blockquote>
		</form>
	<%end if%>
	<h2>Voer gegevens in</h2>
	<form name="form" method="POST" action="?a=import_3">
	<table width="100%" cellpadding=0 cellspacing=2>
	<%for i = 1 to 18%>
	<tr>
		<td><input type=checkbox name="<%=i%>_import" id="<%=i%>_import"></td>
		<td><input type=text size=10 maxlength=10 name="<%=i%>_datum" value="<%=date2datum(date,false)%>"></td>
		<td><input type=text size=3 maxlength=3 name="<%=i%>_bonnr" value=""></td>
		<td>&euro; <input type=text size=10 maxlength=10 name="<%=i%>_bedrag" style="text-align:right" value="" onfocus="document.getElementById('<%=i%>_import').checked = true"/></td>
		<td><input type=text size=30 name="<%=i%>_aan_van" value=""></td>
		<td><input type=text size=40 name="<%=i%>_omschrijving" value=""></td>
		<td><select name="<%=i%>_categorie">
			<option>-</option>
			<%
			set rs = execquery("select naam from categorie WHERE userid="&session("3alansid")&" order by naam asc", adocon)
			do until rs.EOF
				response.write("<option>"&rs("naam")&"</option>")
			rs.movenext : loop
			rs.close : set rs = nothing
			%>
			</select></td>
		<td><select name="<%=i%>_liquide_middel">
			<option>-</option>
			<%
			set rs = execquery("SELECT naam FROM liquide_middel WHERE userid="&session("3alansid")&" ORDER BY naam ASC", adocon)
			do until rs.EOF
				if rs("naam") = "Kas" then
					response.write("<option selected>"&rs("naam")&"</option>")
				else
					response.write("<option>"&rs("naam")&"</option>")
				end if
			rs.movenext : loop
			rs.close : set rs = nothing
			%>
			</select></td>
		<td><input type=text size=8 maxlength=8 name="<%=i%>_afrekeningnr" value=""></td>
	</tr>
	<%next%>
	<tr><td colspan=6 align="center"><br><input type="submit" value="opslaan"></td></tr></table></form>
	<%



elseif request.querystring("a") = "import_2" then
	teller = 0

	Set Upload = Server.CreateObject("Persits.Upload")
	Upload.ProgressID = Request.QueryString("PID")
	Upload.OverwriteFiles = True		'als bestand al bestaat wordt hij hernoemd
	Upload.SetMaxSize 2000000, true		'maximaal 2 Mb
	Count = Upload.Save(Server.MapPath("../../temp/"))

	dim filetype
	filetype = Upload.Form("filetype")

	If Count = 0 Then
		Response.Write "Er zijn geen bestanden geupload."
		Response.End
	End if
	
	Set File = Upload.Files(1)
	%>
	<h2>Importeer gegevens</h2>
	<form action="?a=import_3" name="form" method="POST">
	<table width="100%" cellpadding=0 cellspacing=2>
	<%
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	set file = fs.OpenTextFile(File.Path,1)
	teller = 0
	
	do until file.AtEndOfStream
		
		redim strarray(20)
		x = file.ReadLine
		x = split(x,",")
		teller = teller+1
		
		teller2 = 0
		for each i in x
			teller2 = teller2+1
			strarray(teller2) = i
		next
		
		'strarray is now the array containing all stuff

		dim selectcat, bonnr, bedrag, aanvan, omschrijving, tegenrekening, rekeningnummer, rekeningnaam
		selectcat = ""
		
		
		if filetype = "Rabobank" and teller2>10 then
			datum = strarray(3)
				datum = right(datum,2) & "-" & mid(datum,5,2) & "-" & left(datum,4)
			bonnr = ""
			aanvan = strarray(7)
			omschrijving = trimacols(strarray(11))
			tegenrekening = strarray(6)
			rekeningnummer = trimacols(strarray(1))
			rekeningnaam = ""
			bedrag = strarray(5)
				bedrag = replace(bedrag,".",",")
				if trimacols(strarray(4)) = "D" then
					bedrag = -1 * bedrag
				end if
				bedrag = formatnumber(bedrag)
			if left(omschrijving,12)="Geldautomaat" or instr(omschrijving, "Chipknip") then
				omschrijving = "gepind"
				selectcat = "Kas"
			end if
			if trimacols(strarray(12)) <> "" then
				omschrijving = omschrijving & " - " & trimacols(strarray(12))
			end if
			if left(omschrijving,11)="Pinautomaat" then
				omschrijving = ""
			end if

			if instr(omschrijving,"Rente") then
				selectcat = "Overige inkomsten"
			elseif instr(aanvan,"NOVIB") then
				selectcat = "Overige uitgaven"
			elseif aanvan = "HTM PERS. VERVOER NV" or left(aanvan,2) = "NS" then
				selectcat = "Openbaar vervoer"
			elseif aanvan = "INFORMATIE BEHEER GROEP" then
				selectcat = "Studiefinanciering"
			elseif instr(aanvan,"DYNABYTE") or instr(aanvan,"MEDIA MARKT") or instr(aanvan,"VEVIDA") or instr(aanvan,"Twee Meter Kompjoeters")then
				selectcat = "Computer"
			elseif instr(aanvan,"MOBILE NETHERLANDS") then
				selectcat = "Telefoon"
			elseif instr(aanvan,"AD HOC") then
				selectcat = "Kamerhuur"
			elseif aanvan="STG. HBO HAAGLANDEN" or left(aanvan,4)="HSWS"then
				selectcat = "Bijbaantjes"
			elseif instr(aanvan,"INTERPULSE") then
				selectcat = "Interpulse"
			elseif instr(aanvan,"ALBERT HEIJN") or instr(aanvan,"KONMAR") or instr(aanvan,"C1000") or instr(aanvan,"SNACKBAR HENDO") or instr(aanvan,"DEKAMARKT") or instr(aanvan,"MULTI VLAAI") or instr(aanvan,"SUPER DE BOER") then
				selectcat = "Eten"
			elseif instr(aanvan,"BLOKKER") or instr(aanvan,"GAMMA") or instr(aanvan,"XENOS") or instr(aanvan,"KARWEI") then
				selectcat = "Kamerinrichting"
			elseif instr(aanvan,"Q42") then
				selectcat = "Q42"
			elseif trimacols(tegenrekening) = "3356176560" then
				selectcat = "Spaarrekening"
			elseif trimacols(tegenrekening) = "0335672663" then
				selectcat = "Bank"
			elseif instr(aanvan,"BREAKAWAY CAFE") then
				selectcat = "Uitgaan"
			elseif instr(aanvan, "ZORG EN ZEKERHE") then
				selectcat = "Verzekering"
			elseif instr(aanvan,"HAAGUIT") then
				selectcat = "Vakantie"
			elseif aanvan="GRAND CAFE XIEJE" then
				selectcat = "Uit eten"
			elseif instr(aanvan,"HALFORDS") or instr(aanvan,"FIETSWERELD") or instr(aanvan,"BEVER ZWERFSPORT") then
				selectcat = "Sport"
			elseif (instr(aanvan,"NS-DEN HAAG C.") and instr(strarray(5),"5.8")<>0) or (instr(aanvan,"NS UTRECHT LUN") and instr(strarray(5),"9.8")<>0) or instr(aanvan,"KONMAR 4158") then
				selectcat = "Voorgeschoten"
			elseif instr(aanvan,"NS") then
				selectcat = "Openbaar vervoer"
			end if

		elseif filetype = "Handmade" and teller2>6 then

			datum = strarray(3)
			bonnr = strarray(7)
			aanvan = strarray(4)
			omschrijving = trimacols(strarray(5))
			tegenrekening = ""
			rekeningnummer = ""
			rekeningnaam = strarray(1)
			bedrag = strarray(2)
			
			bedrag = replace(bedrag,".",",")
			bedrag = formatnumber(bedrag)
			
			if left(omschrijving,11)="Pinautomaat" or left(omschrijving,12)="Geldautomaat" or instr(omschrijving, "Chipknip") then
				omschrijving = "gepind"
				selectcat = "Kas"
			end if
			
			if bedrag = "12,50" and omschrijving = "" then
				omschrijving = "Contributie"
			end if
			if instr(lcase(omschrijving),"contributie") then
				selectcat = "contributie"
			elseif instr(lcase(omschrijving),"kleding") and bedrag > 0 then
				selectcat = "kleding in"
			elseif instr(lcase(omschrijving),"kleding") and bedrag < 0 then
				selectcat = "kleding uit"
			elseif instr(lcase(omschrijving),"sponsor") then
				selectcat = "sponsoring"
			elseif instr(lcase(aanvan),"skatebond") then
				selectcat = "skatebond"
			elseif instr(lcase(omschrijving),"website") then
				selectcat = "overige uitgaven"
			end if
		end if



		if checkrekening(datum) and teller2>6 then
			
			response.write("<tr>")
			
			'vinkje
			response.write("<td><input type=checkbox name="""&teller&"_import"" checked></td>")
			
			'datum
			response.write("<td><input type=text size=10 maxlength=10 name="""&teller&"_datum"" value="""& datum &"""></td>")
			
			'bonnummer
			response.write("<td><input type=text size=3 maxlength=3 name="""&teller&"_bonnr"" value=""" & bonnr & """></td>")
			
			'bedrag
			response.write("<td>&euro; <input type=text size=10 maxlength=8 name="""&teller&"_bedrag"" style=""text-align:right"" value="""& replace(bedrag,".","") &"""></td>")
			
			'aan_van
			response.write("<td><input type=text size=30 name="""&teller&"_aan_van"" value="""&trimacols(aanvan)&"""></td>")
			
			'omschrijving
			response.write("<td><input type=text size=40 maxlength=50 name="""&teller&"_omschrijving"" value=""" & omschrijving & """></td>")
			
			'categorie
			response.write("<td><select name="""&teller&"_categorie""><option>-</option>")
			set rs = execquery("select naam from categorie WHERE userid="&session("3alansid")&" order by naam asc", adocon)
			do until rs.EOF
				if lcase(rs("naam")) = lcase(selectcat) then
					response.write("<option selected>"&rs("naam")&"</option>")
				else
					response.write("<option>"&rs("naam")&"</option>")
				end if
			rs.movenext : loop
			rs.close : set rs = nothing
			response.write("</select></td>")
			
			'liquide_middel
			response.write("<td><select name="""&teller&"_liquide_middel""><option>-</option>")
			set rs = execquery("select * from liquide_middel WHERE userid="&session("3alansid")&" order by naam asc", adocon)
			do until rs.EOF
				if rs("rekeningnummer") = rekeningnummer or lcase(rs("naam")) = lcase(rekeningnaam) then
					response.write("<option selected>"&rs("naam")&"</option>")
				else
					response.write("<option>"&rs("naam")&"</option>")
				end if
			rs.movenext : loop
			rs.close : set rs = nothing
			response.write("</select></td>")
			response.write("</tr>")
		end if
	loop
	
	For Each File in Upload.Files
		File.Delete
	Next
	set upload = nothing
	set file = nothing
	set fs = nothing
	%><tr><td colspan=6 align="center"><br><input type="submit" value="opslaan"></td></tr></table></form><%


elseif request.querystring("a") = "import_3" then
	for each i in request.form
		if right(i,6) = "import" and request.form(i) = "on" then
			arrImport = arrImport & i & ";"
		end if
	next
	arrImport = replace(arrImport,"_import","")
	arrImport = split(arrImport,";")
	
	for each i in arrImport
		if not i = "" then
			strQuery = "INSERT INTO overboeking(datum, bedrag, aan_van, omschrijving, categorie, liquide_middel, userid, bonnr, afrekeningnr) VALUES("
			strQuery = strQuery & insertToDb(request.form(i&"_datum"),"date") & ","
			strQuery = strQuery & insertToDb(request.form(i&"_bedrag"),"float") & ","
			strQuery = strQuery & insertToDb(request.form(i&"_aan_van"),"string") & ","
			strQuery = strQuery & insertToDb(left(request.form(i&"_omschrijving"),50),"string") & ","
			strQuery = strQuery & insertToDb(request.form(i&"_categorie"),"string") & ","
			strQuery = strQuery & insertToDb(request.form(i&"_liquide_middel"),"string") & ","
			strQuery = strQuery & insertToDb(session("3alansid"),"int") & ","
			strQuery = strQuery & insertToDb(request.form(i&"_bonnr"),"int") & ","
			strQuery = strQuery & insertToDb(request.form(i&"_afrekeningnr"),"string")
			strQuery = strQuery & ")"
			
			response.write(strQuery & "<br>")
			if not (localhost = true) then
				adocon.execute(strQuery)
			end if
		end if
	next
	%>
	<h2>Import geslaagd</h2>
	De gegevens zijn succesvol in de database geimporteerd.
<%end if

set adocon = nothing
%>
</body>
</html>