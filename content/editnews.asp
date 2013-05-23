<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'></meta>
<title>AKSA News Content Editor</title>
<link rel=STYLESHEET href='..\aksastyle.css' type='text/css'>
<style type='text/css'>
<!--
a:link     { text-decoration:underline; }
-->
</style>
</head>
<h1/><p class='center'>Edit News Content

<form action='newsaction.asp' method='post'>
<table>
<tr>

<%
'on error resume next
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 2
Const adLockOptimistic = 3
Const adUseClient = 3

strSitePath = request.servervariables("PATH_TRANSLATED")
strSitePath = left(strSitePath, len(strSitePath) - (len(strSitePath) - inStrRev(strSitePath, "\")))

session("oTitle") = request("news")

set objConnection = CreateObject("adodb.connection")
objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strSitePath & "\NewsContent.mdb"
set objRecordset = CreateObject("adodb.recordset")
objRecordset.CursorLocation = adUseClient

objRecordset.Open "select * from Items where Title like '" & request("news") & "'", objConnection, adOpenStatic, adLockOptimistic
if objRecordset.EOF then
  objRecordset.Close
  'get the logon and strip off the domain name to get the student's name
  strAuthor = request.ServerVariables("LOGON_USER")
  strAuthor = right(strAuthor, len(strAuthor) - inStr(strAuthor, "\"))
  strAuthor = lcase(strAuthor)
  objRecordset.Open "insert into Items (Title, Author) values ('NewStory', '" & strAuthor & "')"
  objRecordset.Open "select * from Items where Title like 'NewStory'", objConnection, adOpenStatic, adLockOptimistic
end if
response.write "<tr><td><h2/>Title: </td><td><h3/><input type='text' maxlength='50' name='iTitle' size='25' value='" & objRecordset("Title") & "'></td></tr>"
response.write "<tr><td><h2/>Author: </td><td><h3/>" & objRecordset("Author") & "</td></tr>"
response.write "<tr><td><h2/>Posting Date: </td><td><input type='text' maxlength='25' name='iPostingDate' size='25' value='" & objRecordset("PostingDate") & "'></td></tr>"
response.write "<tr><td><h2/>Removal Date: </td><td><input type='text' maxlength='25' name='iRemovalDate' size='25' value='" & objRecordset("RemovalDate") & "'></td></tr>"
response.write "<tr><td><h2/>Sort Order: </td><td><input type='text' maxlength='5' name='iSortOrder' size='5' value='" & objRecordset("SortOrder") & "'></td></tr>"
if objRecordset("Sticky") = True then
  response.write "<tr><td><h2/>Sticky: </td><td><input type='checkbox' name='iSticky' value='True' checked></td></tr>"
else
  response.write "<tr><td><h2/>Sticky: </td><td><input type='checkbox' name='iSticky' value='True'></td></tr>"
end if
response.write "<tr><td><h2/>Headline Image: </td><td><input type='text' maxlength='255' name='iHeadlineImage' size='50' value='" & objRecordset("HeadlineImage") & "'></td></tr>"
response.write "<tr><td><h2/>Headline Image Width: </td><td><input type='text' maxlength='5' name='iHeadlineImageWidth' size='5' value='" & objRecordset("HeadlineImageWidth") & "'></td></tr>"
response.write "<tr><td><h2/>Headline Image Height: </td><td><input type='text' maxlength='5' name='iHeadlineImageHeight' size='5' value='" & objRecordset("HeadlineImageHeight") & "'></td></tr>"
response.write "<tr><td><h2/>Headline Link: </td><td><input type='text' maxlength='255' name='iHeadlineLink' size='50' value='" & objRecordset("HeadlineLink") & "'></td></tr>"
response.write "<tr><td><h2/>Headline: </td><td><input type='text' maxlength='255' name='iHeadline' size='100' value='" & objRecordset("Headline") & "'></td></tr>"
response.write "<tr><td><h2/>Additional Text: </td><td><textarea name='iAdditionalText' cols='80' rows='5'>" & objRecordset("AdditionalText") & "</textarea></td></tr>"
response.write "<tr><td><h2/>Story Image: </td><td><input type='text' maxlength='255' name='iStoryImage' size='50' value='" & objRecordset("StoryImage") & "'></td></tr>"
response.write "<tr><td><h2/>Story Image Width: </td><td><input type='text' maxlength='5' name='iStoryImageWidth' size='5' value='" & objRecordset("StoryImageWidth") & "'></td></tr>"
response.write "<tr><td><h2/>Story Image Height: </td><td><input type='text' maxlength='5' name='iStoryImageHeight' size='5' value='" & objRecordset("StoryImageHeight") & "'></td></tr>"
response.write "<tr><td><h2/>Story: </td><td><textarea name='iStory' cols='80' rows='10'>" & objRecordset("Story") & "</textarea></td></tr>"
response.write "<tr><td><h2/>Hits: </td><td><h3/>" & objRecordset("Hits") & "</td></tr>"
response.write "<tr><td><h2/>Last Hit: </td><td><h3/>" & objRecordset("LastHit") & "</td></tr>"

objRecordset.Close
%>

</table>
<input type='submit' value='Submit' name='submit'>
</form>
</p>

</body>
</html>