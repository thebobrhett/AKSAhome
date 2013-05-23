<%
'****************
'Bob Rhett - Monday, October 6, 2008
'  Created
'****************
'on error resume next

dim objQuality
dim objyield
dim objsm
dim heightfactor

heightfactor = 2

set objQuality = CreateObject("adodb.connection")
objQuality.open = "driver={MySQL ODBC 3.51 Driver};option=16387;server=richmond.aksa.local;user=rootb;password=spandex;DATABASE=quality;"
set objyield = CreateObject("adodb.recordset")
set objsm = CreateObject("adodb.recordset")

response.write "<h1/><p class='center'>Spinning Machine Utilization and Yield<br/>"
response.write "<table width='80%' cellspacing='2'>"
response.write "<tr>"

strSQL = "select * from yield order by sm"
objyield.open strSQL, objQuality

do until objyield.eof
  response.write "<td align='center' valign='bottom' width='7%'>"
  response.write "<h2>" & objyield("yield_pct") & "<br>"
  if cint(objyield("utilization_pct")) > cint(objyield("yield_pct")) then
    response.write "<img src='http://mogsb8/splash/light_gold.gif' alt='" & objyield("utilization_pct") & "% Utilization on SM" & objyield("sm") & "' width='24' height='" & (cint(objyield("utilization_pct")) - cint(objyield("yield_pct"))) * heightfactor & "' border='1'><br>"
    response.write "<img src='http://mogsb8/splash/gold.gif' alt='" & objyield("yield_pct") & "% Yield on SM" & objyield("sm") & "' width='24' height='" & cint(objyield("yield_pct")) * heightfactor & "' border='1'><br>"
  else
    response.write "<img src='http://mogsb8/splash/gold.gif' alt='" & objyield("yield_pct") & "% Yield on SM" & objyield("sm") & "' width='24' height='" & (cint(objyield("yield_pct")) - cint(objyield("utilization_pct"))) * heightfactor & "' border='1'><br>"
    response.write "<img src='http://mogsb8/splash/gold.gif' alt='" & objyield("utilization_pct") & "% Utilization on SM" & objyield("sm") & "' width='24' height='" & cint(objyield("utilization_pct")) * heightfactor & "' border='1'><br>"
  end if
  response.write "</td>"
  objyield.movenext
loop

response.write "</tr>"

objyield.close

strSQL = "select sm from yield order by sm"
objsm.open strSQL, objQuality

response.write "<tr>"

do until objsm.eof
  response.write "<th bgcolor='#955D2C'>"
  response.write "<font color='White' size='2'>SM" & objsm("sm")
  response.write "</th>"
  objsm.movenext
loop

response.write "</tr>"
response.write "</table>"

objsm.close
objQuality.close

set objsm = nothing
set objyield = nothing
set objQuality = nothing

%>
