<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'></meta>
<title>AKSA News Action</title>
</head>

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

'get the logon and strip off the domain name to get the User's name
strUser = request.ServerVariables("LOGON_USER")
strUser = right(strUser, len(strUser) - inStr(strUser, "\"))
strUser = lcase(strUser)

set objConnection = CreateObject("adodb.connection")
objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strSitePath & "NewsContent.mdb"
set objRecordset = CreateObject("adodb.recordset")
objRecordset.CursorLocation = adUseClient

if request("Action") = "Change" or request("Action") = "Copy" or request("Action") = "Delete" or request("Action") = "Preview" or request("Action") = "Reset" then
  oTitle = request("iTitle")
elseif request("Action") = "Approve" or request("Action") = "Reject" then
  oTitle = request("news")
else
  'This is here in case the Title changed coming from the edit page
  oTitle = session("oTitle")
end if

if request("Action") = "Change" then
  objRecordset.Open "select * from Items where Title = '" & oTitle & "'", objConnection, adOpenStatic, adLockOptimistic
  if request("iSticky") = "True" then
    objRecordset("Sticky") = True
  else
    objRecordset("Sticky") = False
  end if
  objRecordset.Update
  objRecordset.Close
end if

if request("Action") = "Preview" then
'  session("scriptdone") = False
  response.redirect "newspreview.asp?news=" & oTitle
'  response.write "newspreview.asp?news=" & oTitle
end if

if request("Action") = "Reset" then
  objRecordset.Open "select * from Items where Title = '" & oTitle & "'", objConnection, adOpenStatic, adLockOptimistic
  objRecordset("Hits") = 0
  objRecordset.Update
  objRecordset.Close
  response.redirect "selectnews.asp"
end if

if request("Action") = "Delete" then
  objRecordset.Open "delete from Items where Title = '" & oTitle & "'", objConnection, adOpenStatic, adLockOptimistic
  'Close up the hole
  objRecordset.Open "select * from Items order by SortOrder", objConnection, adOpenStatic, adLockOptimistic
  intCounter = 1
  do until objRecordset.EOF
    objRecordset("SortOrder") = intCounter
    objRecordset.Update
    objRecordset.MoveNext
    intCounter = intCounter + 1
  loop
  objRecordset.Close
  response.redirect "selectnews.asp"
end if

objRecordset.Open "select * from Items where Title = '" & oTitle & "'", objConnection, adOpenStatic, adLockOptimistic

if request("Action") = "Copy" then
  iTitle = NewStory
  iPostingDate = objRecordset("PostingDate")
  iRemovalDate = objRecordset("RemovalDate")
  iSortOrder = 1
  iHeadlineImage = objRecordset("HeadlineImage")
  iHeadlineImageWidth = objRecordset("HeadlineImageWidth")
  iHeadlineImageHeight = objRecordset("HeadlineImageHeight")
  iHeadlineLink = objRecordset("HeadlineLink")
  iHeadline = objRecordset("Headline")
  iAdditionalText = objRecordset("AdditionalText")
  iStoryImage = objRecordset("StoryImage")
  iStoryImageWidth = objRecordset("StoryImageWidth")
  iStoryImageHeight = objRecordset("StoryImageHeight")
  iStory = objRecordset("Story")
  objRecordset.Close
  objRecordset.Open "insert into Items (Title, Author) values ('NewStory', '" & strUser & "')"
  objRecordset.Open "select * from Items where Title = 'NewStory'", objConnection, adOpenStatic, adLockOptimistic
  objRecordset("PostingDate") = iPostingDate
  objRecordset("RemovalDate") = iRemovalDate
  objRecordset("SortOrder") = iSortOrder
  objRecordset("HeadlineImage") = iHeadlineImage
  objRecordset("HeadlineImageWidth") = iHeadlineImageWidth
  objRecordset("HeadlineImageHeight") = iHeadlineImageHeight
  objRecordset("HeadlineLink") = iHeadlineLink
  objRecordset("Headline") = iHeadline
  objRecordset("AdditionalText") = iAdditionalText
  objRecordset("StoryImage") = iStoryImage
  objRecordset("StoryImageWidth") = iStoryImageWidth
  objRecordset("StoryImageHeight") = iStoryImageHeight
  objRecordset("Story") = iStory
  objRecordset.Update
  'Close up the hole
  objRecordset.Close
  objRecordset.Open "select * from Items order by SortOrder", objConnection, adOpenStatic, adLockOptimistic
  intCounter = 1
  do until objRecordset.EOF
    objRecordset("SortOrder") = intCounter
    objRecordset.Update
    objRecordset.MoveNext
    intCounter = intCounter + 1
  loop
  objRecordset.Close
  response.redirect "selectnews.asp"
elseif request("Action") = "Approve" then
  objRecordset("Test") = False
  objRecordset("Approval") = strUser
  objRecordset.Update
  objRecordset.Close
  response.redirect "selectnews.asp"
elseif request("Action") = "Reject" then
  objRecordset("Test") = True
  objRecordset("Approval") = null
  objRecordset.Update
  objRecordset.Close
  response.redirect "selectnews.asp"
else
  'existing entry
  'Did the Title change?
  if request("iTitle") <> oTitle then
    objRecordset("Title") = request("iTitle")
    objRecordset.Update
    oTitle = request("iTitle")
  end if

  'Is the Sort Order changing?
  if objRecordset("SortOrder") <> cint(request("iSortOrder")) then
    oSortOrder = objRecordset("SortOrder")
    iSortOrder = cint(request("iSortOrder"))
    objRecordset.Close
     'Create a hole
    objRecordset.Open "select * from Items order by SortOrder", objConnection, adOpenStatic, adLockOptimistic
    if oSortOrder > iSortOrder then
      do until objRecordset.EOF
        if objRecordset("SortOrder") => iSortOrder then
          objRecordset("SortOrder") = objRecordset("SortOrder") + 1
          objRecordset.Update
        end if
        objRecordset.MoveNext
      loop
    else
      do until objRecordset.EOF
        if objRecordset("SortOrder") => (iSortOrder + 1) then
          objRecordset("SortOrder") = objRecordset("SortOrder") + 1
          objRecordset.Update
        end if
        objRecordset.MoveNext
      loop
    end if
    objRecordset.Close

    'Change the Sort Order
    objRecordset.Open "select * from Items where Title = '" & request("iTitle") & "'", objConnection, adOpenStatic, adLockOptimistic
    if oSortOrder > iSortOrder then
      objRecordset("SortOrder") = iSortOrder
    else
      objRecordset("SortOrder") = iSortOrder + 1
    end if
    objRecordset.Update
    objRecordset.Close

    'Close up the hole
    objRecordset.Open "select * from Items order by SortOrder", objConnection, adOpenStatic, adLockOptimistic
    intCounter = 1
    do until objRecordset.EOF
      objRecordset("SortOrder") = intCounter
      objRecordset.Update
      objRecordset.MoveNext
      intCounter = intCounter + 1
    loop
    objRecordset.Close

    'Pick up where we left off
    objRecordset.Open "select * from Items where Title = '" & request("iTitle") & "'", objConnection, adOpenStatic, adLockOptimistic
  end if

  if request("Action") <> "Change" then
    if isdate(request("iPostingDate")) then
      objRecordset("PostingDate") = request("iPostingDate")
    else
      objRecordset("PostingDate") = null
    end if
    if isdate(request("iRemovalDate")) then
      objRecordset("RemovalDate") = request("iRemovalDate")
    else
      objRecordset("RemovalDate") = null
    end if
'    objRecordset("SortOrder") = request("iSortOrder")
    if request("iSticky") = "True" then
      objRecordset("Sticky") = True
    else
      objRecordset("Sticky") = False
    end if
    objRecordset("HeadlineImage") = request("iHeadlineImage")
    objRecordset("HeadlineImageWidth") = request("iHeadlineImageWidth")
    objRecordset("HeadlineImageHeight") = request("iHeadlineImageHeight")
    objRecordset("HeadlineLink") = request("iHeadlineLink")
    objRecordset("Headline") = request("iHeadline")
    objRecordset("AdditionalText") = request("iAdditionalText")
    objRecordset("StoryImage") = request("iStoryImage")
    objRecordset("StoryImageWidth") = request("iStoryImageWidth")
    objRecordset("StoryImageHeight") = request("iStoryImageHeight")
    objRecordset("Story") = request("iStory")
    objRecordset.Update
  end if
end if
objRecordset.Close
response.redirect "selectnews.asp"
%>
</body>
</html>