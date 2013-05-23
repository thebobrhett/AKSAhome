<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'></meta>
<title>AKSA Home Page</title>
<link rel=STYLESHEET href='../aksastyle.css' type='text/css'>
</head>

<body link='black' vLink='black'>
<%
'on error resume next

dim objFS
dim objFile

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adUseClient = 3

set objFS = CreateObject("Scripting.FileSystemObject")

strSitePath = request.servervariables("PATH_TRANSLATED")
strSitePath = left(strSitePath, len(strSitePath) - (len(strSitePath) - inStrRev(strSitePath, "\")))

set objConnection = CreateObject("adodb.connection")
objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strSitePath & "NewsContent.mdb"
set objRecordset = CreateObject("adodb.recordset")

objRecordset.Open "select HeadlineScript from Items where Title like '" & request("news") & "'", objConnection, adOpenStatic, adLockOptimistic

set objFile = objFS.OpenTextFile(strSitePath & "headlinescript.inc", ForWriting)
objFile.Write objRecordset("HeadlineScript")
objFile.close

objRecordset.close
objConnection.close

set objFile = nothing
set objFS = nothing
set objRecordset = nothing
set objConnection = nothing

session(request("news") & "headlinescriptdone") = True

response.redirect "http://mogsa4/aksahometestscripting.asp"
'response.write session("backto")
'response.write request.servervariables("http_referer")
'response.redirect request.servervariables("http_referer")
'response.redirect "newspreview.asp?news=" & request("news")
%>
</body>
</html>