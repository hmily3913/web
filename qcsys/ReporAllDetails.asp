<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%

if Instr(session("AdminPurview"),"|40,")=0 then 
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
dim sum1,sum2,sump
sum1=0
sum2=0
sump=0
if showType="UQMD1" then 
%>
 <div id="listtable" style="width:100%; height:100%; overflow:scroll">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验员</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>月份</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>处理笔数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格总笔数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格来料处理率</strong></font></td>
  </tr>
 <%

  dim rs,sql'sql语句
  sql="select * from qcsys_UnQualifiMtrDealSum where UPTdate='"&split(getDateRangebyMonth(laststrmonth&"#"&lastyear),"###")(1)&"'" 
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,0,1
    while(not rs.eof)'填充数据到表格
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("个人")&"</td>"
      Response.Write "<td nowrap>"&laststrmonth&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("本月有效数")&"</td>"
      Response.Write "<td nowrap>"&rs("本月总数")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("本月有效率")&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
			sum1=sum1+rs("本月有效数")
			sum2=sum2+rs("本月总数")
	  rs.movenext
    wend
		sump=sum1*100/sum2
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>"
      Response.Write "<td nowrap></td>" & vbCrLf
      Response.Write "<td nowrap>"&sum1&"</td>"
      Response.Write "<td nowrap>"&sum2&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(sump,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
  rs.close
  set rs=nothing
  %>
  </table>
  </div>
<% 
elseif showType="UQMD2" then 
%>
 <div id="listtable" style="width:100%; height:100%; overflow:scroll">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验员</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>处理笔数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格总笔数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格来料处理率</strong></font></td>
  </tr>
 <%

  sql="select * from qcsys_UnQualifiMtrDealSum where datediff(d,UPTdate,getdate())=1" 
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    while(not rs.eof)'填充数据到表格
	if rs("当日有效率")<>"" then
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("个人")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("UPTdate")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("当日有效数")&"</td>"
      Response.Write "<td nowrap>"&rs("当日总数")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("当日有效率")&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
			sum1=sum1+rs("当日有效数")
			sum2=sum2+rs("当日总数")
	  end if
	  rs.movenext
    wend
		sump=sum1*100/sum2
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>"
      Response.Write "<td nowrap></td>" & vbCrLf
      Response.Write "<td nowrap>"&sum1&"</td>"
      Response.Write "<td nowrap>"&sum2&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(sump,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
  rs.close
  set rs=nothing
  %>
  </table>
  </div>
<% 
elseif showType="UQMD3" then 
%>
 <div id="listtable" style="width:100%; height:100%; overflow:scroll">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验员</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>月份</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>处理笔数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格总笔数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格来料处理率</strong></font></td>
  </tr>
 <%

  sql="select * from qcsys_UnQualifiMtrDealSum where datediff(d,UPTdate,getdate())=1" 
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    while(not rs.eof)'填充数据到表格
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("个人")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&strmonth&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("本月有效数")&"</td>"
      Response.Write "<td nowrap>"&rs("本月总数")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("本月有效率")&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
			sum1=sum1+rs("本月有效数")
			sum2=sum2+rs("本月总数")
	  rs.movenext
    wend
		sump=sum1*100/sum2
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>"
      Response.Write "<td nowrap></td>" & vbCrLf
      Response.Write "<td nowrap>"&sum1&"</td>"
      Response.Write "<td nowrap>"&sum2&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(sump,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
  rs.close
  set rs=nothing
  %>
  </table>
  </div>
<% 
elseif showType="CCP1" then 
%>
 <div id="listtable" style="width:100%; height:100%; overflow:scroll">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验员</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>月份</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验及时笔数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验总笔数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>进料检验及时率</strong></font></td>
  </tr>
 <%
	if Weekday(split(getDateRangebyMonth(strmonth),"###")(0))=1 then
	  sql="select * from qcsys_ComeCheckPromSum where UPTdate=dateadd(d,1,'"&split(getDateRangebyMonth(strmonth),"###")(0)&"')" 
	else  
	  sql="select * from qcsys_ComeCheckPromSum where UPTdate='"&split(getDateRangebyMonth(strmonth),"###")(0)&"'" 
	end if
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    while(not rs.eof)'填充数据到表格
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("个人")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&laststrmonth&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("本月有效数")&"</td>"
      Response.Write "<td nowrap>"&rs("本月总数")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("本月有效率")&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
			sum1=sum1+rs("本月有效数")
			sum2=sum2+rs("本月总数")
	  rs.movenext
    wend
		sump=sum1*100/sum2
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>"
      Response.Write "<td nowrap></td>" & vbCrLf
      Response.Write "<td nowrap>"&sum1&"</td>"
      Response.Write "<td nowrap>"&sum2&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(sump,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
  rs.close
  set rs=nothing
  %>
  </table>
  </div>
<% 
elseif showType="CCP2" then 
%>
 <div id="listtable" style="width:100%; height:100%; overflow:scroll">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验员</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验及时笔数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验总笔数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>进料检验及时率</strong></font></td>
  </tr>
 <%
	if Weekday(dateadd("d",-1,now()))=1 then
	  sql="select * from qcsys_ComeCheckPromSum where datediff(d,UPTdate,getdate())=2" 
	else  
	  sql="select * from qcsys_ComeCheckPromSum where datediff(d,UPTdate,getdate())=1" 
	end if
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    while(not rs.eof)'填充数据到表格
	if rs("当日有效率")<>"" then
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("个人")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("UPTdate")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("当日有效数")&"</td>"
      Response.Write "<td nowrap>"&rs("当日总数")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("当日有效率")&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
			sum1=sum1+rs("当日有效数")
			sum2=sum2+rs("当日总数")
	  end if
	  rs.movenext
    wend
		sump=sum1*100/sum2
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>"
      Response.Write "<td nowrap></td>" & vbCrLf
      Response.Write "<td nowrap>"&sum1&"</td>"
      Response.Write "<td nowrap>"&sum2&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(sump,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
  rs.close
  set rs=nothing
  %>
  </table>
  </div>
<% 
elseif showType="CCP3" then 
%>
 <div id="listtable" style="width:100%; height:100%; overflow:scroll">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验员</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>月份</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验及时笔数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验总笔数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>进料检验及时率</strong></font></td>
  </tr>
 <%
	if Weekday(dateadd("d",-1,now()))=1 then
	  sql="select * from qcsys_ComeCheckPromSum where datediff(d,UPTdate,getdate())=2" 
	else  
	  sql="select * from qcsys_ComeCheckPromSum where datediff(d,UPTdate,getdate())=1" 
	end if
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    while(not rs.eof)'填充数据到表格
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("个人")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&strmonth&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("本月有效数")&"</td>"
      Response.Write "<td nowrap>"&rs("本月总数")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("本月有效率")&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
			sum1=sum1+rs("本月有效数")
			sum2=sum2+rs("本月总数")
	  rs.movenext
    wend
		sump=sum1*100/sum2
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>"
      Response.Write "<td nowrap></td>" & vbCrLf
      Response.Write "<td nowrap>"&sum1&"</td>"
      Response.Write "<td nowrap>"&sum2&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(sump,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
  rs.close
  set rs=nothing
  %>
  </table>
  </div>
<% 
elseif showType="getChart" then 
  sql="select yearnum as 年份,Period as 月份,sum(FinishQualified) as 成品出货检验合格率,sum(UnQualifiMtrDeal) as 不合格来料处理率, "&_
		"sum(ComeCheckA) as 进料检验错误次数,sum(CustomComplain) as 客户投诉次数, "&_
		"sum(FirstCheckAccur) as 首检错误次数,sum(ComeCheckProm) as 进料检验及时率 "&_
		"from ( "&_
		"select yearnum,Period,FinishQualified,UnQualifiMtrDeal, "&_
		"ComeCheckA,CustomComplain,FirstCheckAccur,0 as ComeCheckProm   "&_
		"from qcsys a, "&_
		"(select max(uptdate) as 月末日期 from qcsys  "&_
		"where yearnum=datepart(yy,getdate()) and uptdate is not null  "&_
		"and Period<datepart(mm,getdate()) "&_
		"group by yearnum,Period "&_
		") as b "&_
		"where a.uptdate=b.月末日期 "&_
		"union all "&_
		"select yearnum,Period-1 as Period,0 as FinishQualified,0 as UnQualifiMtrDeal, "&_
		"0 as ComeCheckA,0 as CustomComplain,0 as FirstCheckAccur,ComeCheckProm   "&_
		"from qcsys a, "&_
		"(select min(uptdate) as 月初日期 from qcsys  "&_
		"where yearnum=datepart(yy,getdate()) and uptdate is not null  "&_
		"and Period<=datepart(mm,getdate()) and Period>1 "&_
		"group by yearnum,Period "&_
		") as b "&_
		"where a.uptdate=b.月初日期 "&_
		") as ccc "&_
		"group by yearnum,Period "&_
		"order by yearnum asc,period asc" 
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
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
end if
 %>
