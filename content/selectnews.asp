<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'></meta>
<title>AKSA News Content Selector</title>
<link rel=STYLESHEET href='..\aksastyle.css' type='text/css'>
<style type='text/css'>
<!--
a:link     { color:black; text-decoration:underline; }
a:visited  { color:black; text-decoration:underline; }
-->
</style>
</head>
<h1/><p class='center'>Select News Content to Edit

<table border='1'>
<tr>
<th bgcolor='powderblue'><h2/>Title</th>
<th bgcolor='powderblue'><h2/>Author</th>
<th bgcolor='powderblue'><h2/>Approval</th>
<th bgcolor='powderblue'><h2/>Posting Date</th>
<th bgcolor='powderblue'><h2/>Removal Date</th>
<th bgcolor='powderblue'><h2/>Headline</th>
<th bgcolor='powderblue'><h2/>Hits</th>
<th bgcolor='powderblue'><h2/>Last Hit</th>
<th bgcolor='powderblue'><h2/>Sort Order</th>
<th bgcolor='powderblue'><h2/>Preview</th>
<th bgcolor='powderblue'><h2/>Copy</th>
<th bgcolor='powderblue'><h2/>Delete</th>
</tr>
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

set objConnection = CreateObject("adodb.connection")
objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strSitePath & "\NewsContent.mdb"
set objRecordset = CreateObject("adodb.recordset")
objRecordset.CursorLocation = adUseClient

objRecordset.Open "select * from Items order by SortOrder", objConnection, adOpenStatic, adLockOptimistic

do until objRecordset.EOF
  bPost = "bgcolor='#bbbbbb'"

  if (isnull(objRecordset("PostingDate")) or objRecordset("PostingDate") < Now()) and (isnull(objRecordset("RemovalDate")) or objRecordset("RemovalDate") > Now()) then
    bPost = "bgcolor='#eeeeee'"
  end if

  if len(objRecordset("Title")) = 0 then
    bPost = "bgcolor='#bbbbbb'"
  end if

  if objRecordset("Test") = True then
    bPost = "bgcolor='#bbbbbb'"
  end if

  response.write "<form action='newsaction.asp' method='post'>"
  response.write "<td " & bPost & "><a href='editnews.asp?news=" & objRecordset("Title") & "'><h3/><input type='hidden' name='iTitle' value='" & objRecordset("Title") & "'>" & objRecordset("Title") & "</a></td>"
  response.write "<td " & bPost & "><h3/>" & objRecordset("Author") & "</td>"
  if not isnull(objRecordset("Approval")) then
    response.write "<td " & bPost & "><h3/>" & objRecordset("Approval") & "</td>"
  else
    response.write "<td " & bPost & "><h3/>&nbsp</td>"
  end if
  if isdate(objRecordset("PostingDate")) then
    response.write "<td " & bPost & "><h3/>" & objRecordset("PostingDate") & "</td>"
  else
    response.write "<td " & bPost & "><h3/>&nbsp</td>"
  end if
  if isdate(objRecordset("RemovalDate")) then
    response.write "<td " & bPost & "><h3/>" & objRecordset("RemovalDate") & "</td>"
  else
    response.write "<td " & bPost & "><h3/>&nbsp</td>"
  end if
  response.write "<td " & bPost & "><h3/>" & objRecordset("Headline") & "</td>"
  if objRecordset("Hits") > 0 then
    response.write "<td valign='bottom' " & bPost & "><h3/>" & objRecordset("Hits") & "<br/><input type='submit' name='Action' value='Reset'></td>"
    response.write "<td " & bPost & "><h3/>" & objRecordset("LastHit") & "</td>"
  else
    response.write "<td " & bPost & "><h3/>0</td>"
    response.write "<td " & bPost & "><h3/>&nbsp</td>"
  end if
'  response.write "<td " & bPost & "><input type='text' maxlength='5' name='iSortOrder' size='3' value='" & objRecordset("SortOrder") & "'></td>"
'  response.write "<td " & bPost & "><input type='submit' name='Action' value='Change'></td>"
  response.write "<td valign='bottom' " & bPost & "><input type='text' maxlength='5' name='iSortOrder' size='3' value='" & objRecordset("SortOrder") & "'>"
  if objRecordset("Sticky") = True then
    response.write "<input type='checkbox' name='iSticky' value='True' checked>"
  else
    response.write "<input type='checkbox' name='iSticky' value='True'>"
  end if
  response.write "<br/><input type='submit' name='Action' value='Change'></td>"
  response.write "<td valign='bottom' " & bPost & "><input type='submit' name='Action' value='Preview'></td>"
  response.write "<td valign='bottom' " & bPost & "><input type='submit' name='Action' value='Copy'></td>"
  response.write "<td valign='bottom' " & bPost & "><input type='submit' name='Action' value='Delete'></td>"
  response.write "</tr>"
  objRecordset.MoveNext
  response.write "</form>"
loop
response.write "</table>"
response.write "<p align='center'>"
response.write "<form action='editnews.asp?news=NewStory' method='post'>"
response.write "<input type='submit' name='Action' value='Create New'>"
response.write "</form>"
response.write "</p>"
%>

</p>
</body>
</html>