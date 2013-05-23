<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'></meta>
<title>Full News Story</title>
<link rel=STYLESHEET href='aksastyle.css' type='text/css'>
<script src="myprintln.js"></script>
</head>

<body link='black' vLink='black'>

<div id='logo'>
<a href='http://mogsa4/aksahome.asp'><img src='images\AKGroupLogoSmall.gif' border='0'></a>
<div style='position:absolute; left:10px; top:70px;'><h5><b>A Stretch Above the Rest</b></h5></div>
</div>
</div>

<div id='list'>
<!--#includes file='mainmenu.inc'-->
</div>

<div id='fullstory'>
<%
'****************
'Bob Rhett - Wednesday, October 15, 2008
'  Added ability to insert scripting into a news item.
'****************
'on error resume next
dim strSitePath
dim objConnection
dim objRecordset
'****************
dim objFS
dim objFile
'****************

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adUseClient = 3

strSitePath = request.servervariables("PATH_TRANSLATED")
strSitePath = left(strSitePath, len(strSitePath) - (len(strSitePath) - inStrRev(strSitePath, "\")))

set objConnection = CreateObject("adodb.connection")
objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strSitePath & "Content\NewsContent.mdb"
set objRecordset = CreateObject("adodb.recordset")
objRecordset.CursorLocation = adUseClient

'****************
set objFS = CreateObject("Scripting.FileSystemObject")
'****************

'strSearch = request("news")

objRecordset.Open "select * from Items where Title like '" & request("news") & "'", objConnection, adOpenStatic, adLockOptimistic

'****************
if session("storyscriptdone") = True then
  session("storyscriptdone") = False
else
  if len(objRecordset("StoryScript")) > 0 then
    objRecordset.close
    objConnection.close
    set objRecordset = nothing
    set objConnection = nothing
    set objFS = nothing
    session("backto") = "http://mogsa4/fullstory.asp?news=" & request("news")
    response.redirect "content/storyscript.asp?news=" & request("news")
  end if
end if
'****************

response.write "<br class='clear-left'><div>"
response.write "<h1>" & objRecordset("Headline") & "</h1>"
response.write "<h2>" & objRecordset("AdditionalText") & "</h2>"
'****************
%><!--#include file='content\storyscript.inc'--><%
'****************
if len(objRecordset("StoryImage")) > 0 then
  if objRecordset("StoryImageWidth") = 0 and objRecordset("StoryImageHeight") = 0 then
    response.write "<img src='" & objRecordset("StoryImage") & "' border='0'>"
  elseif objRecordset("StoryImageWidth") > 0 and objRecordset("StoryImageHeight") > 0 then
    response.write "<img src='" & objRecordset("StoryImage") & "' width='" & objRecordset("StoryImageWidth") & "' height='" & objRecordset("StoryImageHeight") & "' border='0'>"
  elseif objRecordset("StoryImageWidth") > 0 then
    response.write "<img src='" & objRecordset("StoryImage") & "' width='" & objRecordset("StoryImageWidth") & "' border='0'>"
  else
    response.write "<img src='" & objRecordset("StoryImage") & "' height='" & objRecordset("StoryImageHeight") & "' border='0'>"
  end if
end if
response.write "<h3>" & replace(objRecordset("Story"), chr(13), "<br/><br/>") & "</h3>"
if isnull(objRecordset("Hits")) then
  objRecordset("Hits") = 1
else
  objRecordset("Hits") = objRecordset("Hits") + 1
end if
objRecordset("LastHit") = Now()
objRecordset.Update
objRecordset.Close
objConnection.close
set objRecordset = nothing
set objConnection = nothing

'****************
set objFile = objFS.GetFile(strSitePath & "content\empty.inc")
objFile.copy strSitePath & "content\storyscript.inc"
'****************
%>
</div>

</body>
</html>