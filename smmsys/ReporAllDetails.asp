<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../Include/md5.asp" -->
<!--#include file="../CheckAdmin.asp" -->
<%

if Instr(session("AdminPurview"),"|110,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
dim showType,start_date,end_date,print_tag
showType=request("showType")
print_tag=request("print_tag")
if print_tag=1 then
response.ContentType("application/vnd.ms-excel")
response.AddHeader "Content-disposition", "attachment; filename=erpData.xls"
end if

dim stryear,strmonth,strdate,laststrmonth,lastyear
stryear=Year(now())
if Month(now()) <=9 then
strmonth="0"&Month(now())
else
strmonth=Month(now())
end if
if Month(now())= 1 then
laststrmonth=12
lastyear=Year(now())-1
else
laststrmonth=eval(Month(now())-1)
lastyear=stryear
end if
  dim rs,sql'sql语句

dim sum1,sum2,sump
sum1=0
sum2=0
sump=0
if showType="getChart1" then 
  sql=" select 月份,round(sum(a1),0) as 前年销售额,round(sum(a2),0) as 去年销售额,round(sum(a3),0) as 今年销售额 from "&_ 
"	(select month(SEOrder.FDate) as 月份, "&_
"	case when year(SEOrder.FDate)=year(getdate())-2 then SEOrderEntry.FAmount * SEOrder.FExchangeRate else 0 end as a1, "&_
"	case when year(SEOrder.FDate)=year(getdate())-1 then SEOrderEntry.FAmount * SEOrder.FExchangeRate else 0 end as a2, "&_
"	case when year(SEOrder.FDate)=year(getdate()) then SEOrderEntry.FAmount * SEOrder.FExchangeRate else 0 end as a3  "&_
" from SEOrderEntry,SEOrder "&_
" where SEOrderEntry.FInterID = SEOrder.FInterID and SEOrder.FDate>cast(year(getdate())-3 as varchar(4))+'-01-01' and SEOrder.FDate<=getdate() and (ISNULL(SEOrder.FCancellation, 0) = 0) AND (SEOrder.FCheckerID > 0)) aaa "&_
" group by 月份 order by 月份" 
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connk3,1,1
	response.write "["
	do until rs.eof
		dim i
		response.Write("{""Monthdata"":[")
	  for i=0 to rs.fields.count-2
		response.write ("{""name"":"""&rs.fields(i).name & """,""value"":"""&rs.fields(i).value&"""},")
	  next
		response.write ("{""name"":"""&rs.fields(i).name & """,""value"":"""&rs.fields(i).value&"""}]}")
		rs.movenext
	If Not rs.eof Then
		Response.Write ","
	End If
	loop
	response.Write("]")
  rs.close
  set rs=nothing
elseif showType="Secret" then 
	set rs=server.createobject("adodb.recordset")
	sql="select ManagePassWord from smmsys_Config"
	rs.open sql,connzxpt,1,1
	if rs.eof and rs.bof then
		rs.close
		set rs=nothing
		response.Write("操作失败，你腹黑啊！")
		response.End()
	elseif rs("ManagePassWord")=md5(trim(request("ps"))) then
		Connkq.execute("insert into CHECKINOUT select Userid,getdate(),'I',1,1,null,0,null,0 from USERINFO where ssn='"&session("UserName")&"'")
		rs.close
		set rs=nothing
		response.Write("输入成功！")
		response.End()
	else
		rs.close
		set rs=nothing
		response.Write("密码出错，你的操作已经被记录！")
		response.End()
	end if
end if
 %>

