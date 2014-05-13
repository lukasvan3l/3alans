<%
dim rs, rs2, rs3, i, maand, intxt, uittxt, jaar, categorie, liqmid, file
dim totaaluit, totaalin, euriuit, euriin, datum, saldo, fs, fo, x, teller, teller2, programName
dim arrImport, strQuery, count, Upload

programName = "3aLans v .007"

dim localhost
localhost = false
if request.servervariables("HTTP_HOST") = "localhost" then
	localhost = true
end if

dim totaal
totaal = 0

if (session("3alansname") = "" OR session("3alansid") = "") AND request.servervariables("SCRIPT_NAME") <> "/3alans/login.asp" then
	dim gotoUrl
	gotoUrl = "http://"
	gotoUrl = gotoUrl & request.servervariables("HTTP_HOST")
	gotoUrl = gotoUrl & request.servervariables("SCRIPT_NAME")
	gotoUrl = gotoUrl & "?from=login"
	for each i in request.querystring()
		gotoUrl = gotoUrl & "&" & i & "=" & request.querystring(i)
	next
	response.redirect("/3alans/login.asp?goto=" & gotoUrl)
	response.end
end if

Sub maandvanjaar(maandnr)
	Select case maandnr
		case "1"  response.write("jan")
		case "2"  response.write("feb")
		case "3"  response.write("mrt")
		case "4"  response.write("apr")
		case "5"  response.write("mei")
		case "6"  response.write("jun")
		case "7"  response.write("jul")
		case "8"  response.write("aug")
		case "9"  response.write("sep")
		case "10"  response.write("okt")
		case "11"  response.write("nov")
		case "12"  response.write("dec")
	End Select
End Sub

Function ExecQuery (strQ, objCon)
	Dim objRS
	
	Set objRS = Server.CreateObject("ADODB.RecordSet")
	
	objRS.CursorLocation = adUseClient
	objRS.CursorType = adOpenForwardOnly
	objRs.LockType = adLockReadOnly
	
	' Response.Write strQ & "<BR>"
	
	objRS.Open strQ, objCon, , , adCmdText
	Set objRS.ActiveConnection = Nothing
	
	Set ExecQuery = objRS
End Function

Function Date2Datum(input, tijd)
	if input <> "" then
		Date2Datum = tweegetallen(day(input)) & "-"
		Date2Datum = Date2Datum & tweegetallen(month(input)) & "-"
		Date2Datum = Date2Datum & year(input)
		if tijd = true then
			Date2Datum = Date2Datum & " " & tweegetallen(hour(input)) & ":"
			Date2Datum = Date2Datum & tweegetallen(minute(input)) & ":"
			Date2Datum = Date2Datum & tweegetallen(second(input))
		end if
	end if
End Function

Function replaceHTMLTags(input)
	'laat <strong> zien ipv dat de tekst vetgedrukt wordt
 	'replaceHTMLTags = changeQuotes(input)
	replaceHTMLTags = Replace(input, """","&quot;")
	replaceHTMLTags = Replace(replaceHTMLTags, "'", "&acute;")
	replaceHTMLTags = Replace(replaceHTMLTags, "<", "&lt;")
	replaceHTMLTags = Replace(replaceHTMLTags, ">", "&gt;")
	replaceHTMLTags = Replace(replaceHTMLTags, "€", "&euro;")
End Function

Function RandomNumber(Aantal)
	Randomize
	RandomNumber = Int(Aantal * Rnd) + 1
End Function

function tweegetallen(input)
	if len(input) = 1 then
		tweegetallen = Cstr("0") & Cstr(input)
	else
		tweegetallen = Cstr(input)
	end if
end function

function changeQuotes(input)
	changeQuotes = replace(input, "'", "''")
	changeQuotes = trim(changeQuotes)
end function

Sub writePageCountLink(link, tekst, qstr)
	Response.Write(	"<A HREF=""" &request.servervariables("SCRIPT_NAME")& "?pagina=" & link )
	for each i in qstr
		if not request.querystring(i) = "" then
			response.write( "&" & i & "=" &request.querystring(i) )
		end if
	next
	response.write( """" )
	response.write( " title=""Naar pagina " )
	response.write( link & """>" & tekst & "</A> " )
End Sub

Sub Print_Navigation(newsRS, nPage, Geschied)
	Dim nRecCount	' Number of records found
	Dim nPageCount	' Number of pages of records we have
	Dim lokatie     ' Geeft aan waar vandaan deze sub is aangeroepen (blok of content)
	Dim p
	
	nRecCount = newsRS.RecordCount
	nPageCount = newsRS.PageCount
	
	'qstr bevat alle querystrings die meegegeven worden als je op volgende of een nummertje klikt.
	'om zelf eentje toe te voegen, verhoog qstr(int) met ééntje, en voeg je eigen querystring toe
	'aan de array hieronder.
	Dim qstr(12)		' Mee te geven querystrings
	qstr(0) = "id"
	qstr(1) = "m"
	qstr(2) = "r_id"
	qstr(3) = "m_id"
	qstr(4) = "action"
	qstr(5) = "actie"
	qstr(6) = "zoek"
	qstr(7) = "f_id"
	qstr(8) = "keywords"
	qstr(9) = "l_id"
	qstr(10) = "sub"
	qstr(11) = "c_id"
	qstr(12) = "orderby"
	
	If nPage < 1 Or nPage > nPageCount Then
		nPage = 1
	End If
	
	dim vorigetekst, volgendetekst
	vorigetekst = "vorige"
	volgendetekst = "volgende"
	
	If Not (nPageCount = 1) Then
		
		if nPageCount < (2*geschied)+1 then
			If Not nPage = 1 Then
				call writePageCountLink(nPage-1, vorigetekst, qstr)
			else
				response.write(vorigetekst)
			End If
			
			response.write(" | ")
			
			For p = 1 To nPageCount
				If Not Cint(nPage) = Cint(p) then
					if Cint(nPageCount) > 9 then
						call writePageCountLink(p,tweegetallen(p),qstr)
					else
						call writePageCountLink(p,p,qstr)
					end if
				else
					if Cint(nPageCount) > 9 then
						response.write( "<strong><u>" & tweegetallen(p) & "</u></strong> " )
					else
						response.write( "<strong><u>" & p & "</u></strong> " )
					end if
				end if
			Next
			
			response.write("| ")
			
			If Not Cint(nPage) = Cint(nPageCount) Then
				call writePageCountLink(nPage+1, volgendetekst, qstr)
			else
				response.write(volgendetekst)
			End If
		
		else
		
			If Not nPage = 1 Then
				call writePageCountLink(1, "<<", qstr)
				response.write(" ")
				call writePageCountLink(nPage-1, "<", qstr)
			else
				response.write("<< <")
			End If
			
			response.write(" | ")
			
			if npage < geschied+1 then
				For p = 1 to (2*geschied)+1
					If Not Cint(nPage) = Cint(p) then
						if Cint(nPageCount) > 9 then
							call writePageCountLink(p,tweegetallen(p),qstr)
						else
							call writePageCountLink(p,p,qstr)
						end if
					else
						if Cint(nPageCount) > 9 then
							response.write( "<strong><u>" & tweegetallen(p) & "</u></strong> " )
						else
							response.write( "<strong><u>" & p & "</u></strong> " )
						end if
					end if
				next
			elseif npage > nPageCount-geschied then
				For p = nPageCount - (2*geschied) to nPageCount
					If Not Cint(nPage) = Cint(p) then
						if Cint(nPageCount) > 9 then
							call writePageCountLink(p,tweegetallen(p),qstr)
						else
							call writePageCountLink(p,p,qstr)
						end if
					else
						if Cint(nPageCount) > 9 then
							response.write( "<strong><u>" & tweegetallen(p) & "</u></strong> " )
						else
							response.write( "<strong><u>" & p & "</u></strong> " )
						end if
					end if
				next
			else
				For p = 1 to nPageCount
					if (p >= (nPage-geschied) AND p <= (nPage+geschied)) then
						If Not Cint(nPage) = Cint(p) then
							if Cint(nPageCount) > 9 then
								call writePageCountLink(p,tweegetallen(p),qstr)
							else
								call writePageCountLink(p,p,qstr)
							end if
						else
							if Cint(nPageCount) > 9 then
								response.write( "<strong><u>" & tweegetallen(p) & "</u></strong> " )
							else
								response.write( "<strong><u>" & p & "</u></strong> " )
							end if
						end if
					end if
				Next
			end if
			
			response.write("| ")
			
			If Not Cint(nPage) = Cint(nPageCount) Then
				call writePageCountLink(nPage+1, ">", qstr)
				response.write(" ")
				call writePageCountLink(nPageCount, ">>", qstr)
			else
				response.write("> ")
				response.write(">>")
			End If
		end if
	End If
End Sub

function trimacols(val)
	if val = "" or isnumeric(val) or isnull(val) then
		val = val
	else
		if left(val,1) = """" then
			val = right(val,len(val)-1)
		end if
		
		if right(val,1) = """" then
			val = left(val,len(val)-1)
		end if
	end if
	trimacols = val
end function

function checkrekening(datum)
	checkrekening = false
	
	if not datum = "" and not datum = "--" then
		checkrekening = true
		
		if trimacols(strarray(6)) = "0335672663" then		'=rekening_aan_van, moet uit db gehaald worden!!
			checkrekening = false
		end if
	end if
end function

function saldoOpDag(datum)
	set rs = adocon.execute("SELECT sum(bedrag) AS startsaldo FROM overboeking WHERE categorie='Beginsaldo' and userid="&session("3alansid"))
	saldoOpDag = rs("startsaldo")
	rs.close : set rs = nothing
	
	set rs = adocon.execute("SELECT sum(bedrag) AS saldo FROM overboeking WHERE not (categorie='Kas' OR categorie='Spaarrekening' OR categorie='Beginsaldo' OR categorie='Bank') AND datum<=FORMAT('" & datum & "','dd-mm-yyyy') and userid="&session("3alansid"))
	if not isnull(rs("saldo")) then
		saldoOpDag = saldoOpDag + rs("saldo")
	end if
	rs.close : set rs = nothing
end function

function insertToDb(txtValue, txtType)
	dim result
	if trim(txtValue) = "" then
		result = "null"
	elseif txtType = "int" then
		if isnumeric(txtValue) then
			result = txtValue
		else
			result = "null"
		end if
	elseif txtType = "float" then
		result = replace(txtValue,".",".")
		result = replace(result,",",".")
	elseif txtType = "string" then
		result = "'" & replace(txtValue,"'","''") & "'"
	elseif txtType = "bool" then
		if txtValue = "on" then
			result = "true"
		else
			result = "false"
		end if
	elseif txtType = "date" then
		result = "FORMAT('" & txtValue &"','dd-mm-yyyy')"
	end if
	
	insertToDb = result
end function
%>