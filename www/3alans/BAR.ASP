<%@EnableSessionState=False%>

<%
Response.Expires = -1
PID = Request("PID")
TimeO = Request("to")

Set UploadProgress = Server.CreateObject("Persits.UploadProgress")
format = "%TUploading files...%t%B3%T%R left (at %S/sec) %r%U/%V(%P)%l%t"
bar_content = UploadProgress.FormatProgress(PID, TimeO, "#00007F", format)

If "" = bar_content Then
	%>
	<HTML>
	<HEAD>
	<TITLE>Upload Finished</TITLE>
	<SCRIPT LANGUAGE="JavaScript">
	function CloseMe()
	{
		window.parent.close();
		return true;
	}
	</SCRIPT>
	</HEAD>
	<BODY OnLoad="CloseMe()">
	</BODY>
	</HTML>
	<%
Else    ' Not finished yet
	%>
	<HTML>
	<HEAD>
	
	<!--%  If left(bar_content, 1) <> "." Then %-->
	
	<meta HTTP-EQUIV="Refresh" CONTENT="1;URL=<%
	 Response.Write Request.ServerVariables("URL")
	 Response.Write "?to=" & TimeO & "&PID=" & PID %>">
	
	<!--% End If %-->
	
	<TITLE>Uploading Files...</TITLE>
	
	<style type='text/css'>td {font-family:arial; font-size: 9pt } td.spread {font-size: 6pt; line-height:6pt } td.brick {font-size:6pt; height:12px}</style>
	
	</HEAD>
	<BODY BGCOLOR="#C0C0B0" topmargin=0>
	<% = bar_content %>
	</BODY>
	</HTML>

<% End If %>