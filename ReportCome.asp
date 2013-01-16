<%
'分流跳转页面


dim subid
subid=request.QueryString("sub")
select case subid
case 1
if Instr(session("AdminPurview"),"|10,")>0 then
response.redirect "salesys/ReporAll.asp"
end if
case 2
if Instr(session("AdminPurview"),"|20,")>0 then
response.redirect "purchasesys/ReporAll.asp"
end if
case 3
if Instr(session("AdminPurview"),"|30,")>0 then
response.redirect "manusys/ReporAll.asp"
end if
case 4
if Instr(session("AdminPurview"),"|40,")>0 then
response.redirect "qcsys/ReporAll.asp"
end if
case 5
if Instr(session("AdminPurview"),"|50,")>0 then
response.redirect "stocksys/ReporAll.asp"
end if
case 6
if Instr(session("AdminPurview"),"|60,")>0 then
response.redirect "technologysys/ReporAll.asp"
end if
case 7
if Instr(session("AdminPurview"),"|70,")>0 then
response.redirect "engineersys/ReporAll.asp"
end if
case 8
if Instr(session("AdminPurview"),"|80,")>0 then
response.redirect "financesys/ReporAll.asp"
end if
case 9
if Instr(session("AdminPurview"),"|90,")>0 then
response.redirect "hrsys/ReporAll.asp"
end if
case 10
if Instr(session("AdminPurview"),"|100,")>0 then
response.redirect "managesys/ReporAll.asp"
end if
end select
%>