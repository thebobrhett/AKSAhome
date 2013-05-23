'*************
'Bob Rhett - Tuesday, October 27,2009	Created.
'  This program replaces the day of the week Subway news item with one that is generic.
'
'Bob Rhett - Tuesday, February 8, 2011
'  Copied to use for Garden Fresh Deli
'*************
'on error resume next

Const adOpenStatic = 2
Const adOpenForwardOnly = 0
Const adLockOptimistic = 3
Const adUseClient = 3

dim objDBNews
dim objRSNews
dim strSQL

set objDBNews = CreateObject("adodb.connection")
objDBNews.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\inetpub\wwwroot\content\NewsContent.mdb"
set objRSNews = CreateObject("adodb.recordset")
objRSNews.CursorLocation = adUseClient

strSQL = "select * from Items where Title='Deli" & WeekdayName(Weekday(now)) & "'"
objRSNews.Open strSQL, objDBNews, adOpenStatic, adLockOptimistic
if not objRSNews.eof then
  'Check to see if this is currently posted.
  if (isnull(objRSNews("PostingDate")) or objRSNews("PostingDate") < Now()) and (isnull(objRSNews("RemovalDate")) or objRSNews("RemovalDate") > Now()) then
    objRSNews.close
    strSQL = "update Items set Test=True where Title='Deli" & WeekdayName(Weekday(now)) & "'"
    objRSNews.open strSQL, objDBNews, adOpenStatic, adLockOptimistic
    strSQL = "update Items set Test=False where Title='DeliNews'"
    objRSNews.open strSQL, objDBNews, adOpenStatic, adLockOptimistic
  end if
end if

objDBNews.close
set objDBNews = nothing
set objRSNews = nothing
