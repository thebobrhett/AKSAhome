<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'></meta>
<title>News Story Preview</title>
<link rel=STYLESHEET href='../aksastyle.css' type='text/css'>
</head>

<body link='black' vLink='black'>

<div id='logo'>
<a href='http://mogsa4/newhome.asp'><img src='..\images\AKGroupLogoSmall.gif' border='0'></a>
<div style='position:absolute; left:10px; top:70px;'><h5><b>A Stretch Above the Rest</b></h5></div>
</div>
</div>

<div id='content-rotator'>
<h1/><font color='red'><br/>Website Content Preview Page</font>
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
'dim objFS
'dim objFile
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
objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strSitePath & "NewsContent.mdb"
set objRecordset = CreateObject("adodb.recordset")
objRecordset.CursorLocation = adUseClient

'****************
'set objFS = CreateObject("Scripting.FileSystemObject")
'****************

'strSearch = request("news")

objRecordset.Open "select * from Items where Title like '" & request("news") & "'", objConnection, adOpenStatic, adLockOptimistic

'****************
'if session("previewscriptdone") = True then
'  session("previewscriptdone") = False
'else
'  if len(objRecordset("HeadlineScript")) > 0 or len(objRecordset("StoryScript")) > 0 then
'    objRecordset.close
'    objConnection.close
'    set objRecordset = nothing
'    set objConnection = nothing
'    set objFS = nothing
'    session("backto") = "http://mogsa4/content/newspreview.asp?news=" & request("news")
'    response.redirect "previewscript.asp?news=" & request("news")
'  end if
'end if
'****************

response.write "<h1/><font color='red'><u>HEADLINE</u></font><br/>"

response.write "<br class='clear-left'><div>"
if len(objRecordset("HeadlineLink")) = 0 or isnull(objRecordset("HeadlineLink")) then
'  if len(objRecordset("HeadlineLink")) > 0 then
    if objRecordset("HeadlineImageWidth") = 0 and objRecordset("HeadlineImageHeight") = 0 then
      response.write "<p class='align-left'><img src='" & objRecordset("HeadlineImage") & "' border='0'></p>"
    elseif objRecordset("HeadlineImageWidth") > 0 and objRecordset("HeadlineImageHeight") > 0 then
      response.write "<p class='align-left'><img src='" & objRecordset("HeadlineImage") & "' width='" & objRecordset("HeadlineImageWidth") & "' height='" & objRecordset("HeadlineImageHeight") & "' border='0'></p>"
    elseif objRecordset("HeadlineImageWidth") > 0 then
      response.write "<p class='align-left'><img src='" & objRecordset("HeadlineImage") & "' width='" & objRecordset("HeadlineImageWidth") & "' border='0'></p>"
    else
      response.write "<p class='align-left'><img src='" & objRecordset("HeadlineImage") & "' height='" & objRecordset("HeadlineImageHeight") & "' border='0'></p>"
    end if
'  end if
else
  if len(objRecordset("HeadlineImage")) > 0 then
    if objRecordset("HeadlineImageWidth") = 0 and objRecordset("HeadlineImageHeight") = 0 then
      response.write "<p class='align-left'><a href='" & objRecordset("HeadlineLink") & "'><img src='" & objRecordset("HeadlineImage") & "' border='0'></a></p>"
    elseif objRecordset("HeadlineImageWidth") > 0 and objRecordset("HeadlineImageHeight") > 0 then
      response.write "<p class='align-left'><a href='" & objRecordset("HeadlineLink") & "'><img src='" & objRecordset("HeadlineImage") & "' width='" & objRecordset("HeadlineImageWidth") & "' height='" & objRecordset("HeadlineImageHeight") & "' border='0'></a></p>"
    elseif objRecordset("HeadlineImageWidth") > 0 then
      response.write "<p class='align-left'><a href='" & objRecordset("HeadlineLink") & "'><img src='" & objRecordset("HeadlineImage") & "' width='" & objRecordset("HeadlineImageWidth") & "' border='0'></a></p>"
    else
      response.write "<p class='align-left'><a href='" & objRecordset("HeadlineLink") & "'><img src='" & objRecordset("HeadlineImage") & "' height='" & objRecordset("HeadlineImageHeight") & "' border='0'></a></p>"
    end if
  end if
end if
if len(objRecordset("HeadlineLink")) = 0 or isnull(objRecordset("HeadlineLink")) then
  response.write "<h1>" & objRecordset("Headline") & "</h1>"
else
  response.write "<a href='" & objRecordset("HeadlineLink") & "'><h1>" & objRecordset("Headline") & "</h1></a>"
end if
'****************
  %><!--#include file='headlinescript.inc'--><%
'****************
response.write "<h2>" & objRecordset("AdditionalText") & "</h2>"
if len(objRecordset("Story")) = 0 or isnull(objRecordset("Story")) then
  response.write "</div><br class='clear-left'><hr/>"
else
  response.write "<a href='http://mogsa4/fullstory.asp?news=" & objRecordset("Title") & "'><h4/>Read more about this...</a></div><br class='clear-left'><hr/>"
end if

response.write "<h1/><font color='red'><u>FULL STORY</u></font><br/>"

if not isnull(objRecordset("Story")) then
  response.write "<br class='clear-left'><div>"
  response.write "<h1>" & objRecordset("Headline") & "</h1>"
  response.write "<h2>" & objRecordset("AdditionalText") & "</h2>"
'****************
  %><!--#include file='storyscript.inc'--><%
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
else
  response.write "<h1>* Nothing to Display *</h1>"
end if
objRecordset.close
objConnection.close
set objRecordset = nothing
set objConnection = nothing

'****************
'set objFile = objFS.GetFile(strSitePath & "empty.inc")
'objFile.copy strSitePath & "headlinescript.inc"
'objFile.copy strSitePath & "storyscript.inc"
'****************

response.write "<form action='newsaction.asp?news=" & request("news") & "' method='post'>"
%>
<table>
<tr>
<td><input type='submit' name='Action' value='Approve'></td>
<td><input type='submit' name='Action' value='Reject'></td>
</tr>
</table>
</form>
</div>

</body>
</html>