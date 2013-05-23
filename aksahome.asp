<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'></meta>
<title>AKSA Home Page</title>
<link rel=STYLESHEET href='http://mogsa4/aksastyle.css' type='text/css'>
<script src='myprintln.js'></script>
</head>

<body link='black' vLink='black'>

<div id='logo'>
<img src='images\AKGroupLogoSmall.gif'>
<div style='position:absolute; left:10px; top:70px;'><h5><b>A Stretch Above the Rest</b></h5></div>
</div>

<%
Set FS=Server.CreateObject("Scripting.FileSystemObject")
Set RS=FS.OpenTextFile("d:\Scripts\homecounter.txt", ForReading, False)
fcount=RS.ReadLine
RS.Close

fcount=fcount+1

'This code is disabled due to the write access security on our server:
Set RS=FS.OpenTextFile("d:\Scripts\homecounter.txt", ForWriting, False)
RS.Write fcount
RS.Close

Set RS=Nothing
Set FS=Nothing
%>

<div id='list'>
<!--#includes file='mainmenu.inc'-->
<h5/>This page has been visited <%=fcount%>  times.
</div>

<div id='content-rotator'>
<%
set cr=server.createobject("MSWC.ContentRotator")
response.write (cr.ChooseContent("content/HeaderContent.txt"))
'response.write (cr.ChooseContent("content/Test_Pledges.txt"))
%>
</div>

<div id='weather'>
<a href='http://www.weatherforyou.com/weather/south+carolina/radar/latest.php'><img src="http://www.weatherforyou.net/fcgi-bin/hw3/hw3.cgi?config=png&forecast=metar&alt=hwihourly&place=goose+creek&state=sc&zipcode=29445&country=us&county=45015&zone=scz045&icao=KCHS&hwvbg=dddddd&hwvtc=black&metric=0&hwvdisplay=Current+Conditions"/></a>
</div>

<div id='news'>
<%
on error resume next
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 2
Const adLockOptimistic = 3
Const adUseClient = 3

strSitePath = request.servervariables("PATH_TRANSLATED")
strSitePath = left(strSitePath, len(strSitePath) - (len(strSitePath) - inStrRev(strSitePath, "\")))

set objConnection = CreateObject("adodb.connection")
objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strSitePath & "Content\NewsContent.mdb"
set objRecordset = CreateObject("adodb.recordset")
objRecordset.CursorLocation = adUseClient

objRecordset.Open "select * from Items order by SortOrder", objConnection, adOpenStatic, adLockOptimistic

do until objRecordset.EOF
  if (isnull(objRecordset("PostingDate")) or objRecordset("PostingDate") < Now()) and (isnull(objRecordset("RemovalDate")) or objRecordset("RemovalDate") > Now()) then
    bPost = True
  else
    bPost = False
  end if

  if len(objRecordset("Title")) = 0 then bPost = False
  if objRecordset("Test") = True then bPost = False

  if bPost = True then
    response.write "<br class='clear-left'><div>"
    if len(objRecordset("HeadlineLink")) = 0 or isnull(objRecordset("HeadlineLink")) then
      if len(objRecordset("HeadlineImage")) > 0 then
        if objRecordset("HeadlineImageWidth") = 0 and objRecordset("HeadlineImageHeight") = 0 then
          response.write "<p class='align-left'><img src='" & objRecordset("HeadlineImage") & "' border='0'></p>"
        elseif objRecordset("HeadlineImageWidth") > 0 and objRecordset("HeadlineImageHeight") > 0 then
          response.write "<p class='align-left'><img src='" & objRecordset("HeadlineImage") & "' width='" & objRecordset("HeadlineImageWidth") & "' height='" & objRecordset("HeadlineImageHeight") & "' border='0'></p>"
        elseif objRecordset("HeadlineImageWidth") > 0 then
          response.write "<p class='align-left'><img src='" & objRecordset("HeadlineImage") & "' width='" & objRecordset("HeadlineImageWidth") & "' border='0'></p>"
        else
          response.write "<p class='align-left'><img src='" & objRecordset("HeadlineImage") & "' height='" & objRecordset("HeadlineImageHeight") & "' border='0'></p>"
        end if
      end if
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
      response.write "<h1/>" & objRecordset("Headline")
    else
      response.write "<a href='" & objRecordset("HeadlineLink") & "'><h1/>" & objRecordset("Headline") & "</a>"
    end if
    response.write "<h2/>" & objRecordset("AdditionalText")
    if len(objRecordset("Story")) = 0 or isnull(objRecordset("Story")) then
      response.write "</div><br class='clear-left'><hr/>"
    else
      response.write "<a href='fullstory.asp?news=" & objRecordset("Title") & "'><h4/>Read more about this...</a></div><br class='clear-left'><hr/>"
    end if
  end if
  objRecordset.MoveNext
loop
%>
</div>

<!--#includes file='mousetail.inc'-->

</body>
</html>