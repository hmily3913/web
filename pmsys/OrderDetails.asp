<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurviewFLW"),"|210,")=0 then 
'  response.write "{""page"":""1"",""total"":""1"",""rows"":[{""id"":""0"",""cell"":["
'  response.write ("""<font color='red')>你不具有该管理模块的操作权限，请返回！</font>""]}]}")
'  response.end
'end if
'========判断是否具有管理权限
dim showType,start_date,end_date,print_tag,UserName,AdminName
UserName=session("UserName")
AdminName=session("AdminName")
showType=request("showType")
print_tag=request("print_tag")
if print_tag=1 then
response.ContentType("application/vnd.ms-excel")
response.AddHeader "Content-disposition", "attachment; filename=erpData.xls"
end if
if showType="DetailsList" then 
	Dim rs,sql
	Dim tempStr1:tempStr1="临时字符串"
	Dim RowsNum:RowsNum=-1
	dim lastweek,daycount,lastday
	sql="select max(date) as lastday from Calendar a where DATENAME(week,getdate())+4=a.weeknum and datediff(yy,a.date,getdate())=0"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
	lastday=rs("lastday")
	dim TotalArr(8)
%>
<div id="listtable" style="width:100%; height:100%; overflow:scroll">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
<tr>
	<td class="tablemenu" colspan="35" height="20" width="100%" id="formove"><font color="#15428B"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:$('#listDiv').hide();$('#QueryTable').show();" >&nbsp;<strong>生管订单交期分布表</strong></font></td>
</tr>
<tr height="12" align="center">
	<td nowrap rowspan="2" bgcolor="#8DB5E9"><strong>序号</strong></td>
	<td nowrap rowspan="2" bgcolor="#8DB5E9"><strong>销售订单号</strong></td>
	<td nowrap rowspan="2" bgcolor="#8DB5E9"><strong>生产任务单</strong></td>
	<td nowrap rowspan="2" bgcolor="#8DB5E9"><strong>产品名称</strong></td>
	<td nowrap rowspan="2" bgcolor="#8DB5E9"><strong>产品编号</strong></td>
	<td nowrap rowspan="2" bgcolor="#8DB5E9"><strong>产品分类</strong></td>
	<td nowrap rowspan="2" bgcolor="#8DB5E9"><strong>生产部门</strong></td>
	<td nowrap rowspan="2" bgcolor="#8DB5E9"><strong>订单数量</strong></td>
	<td nowrap rowspan="2" bgcolor="#8DB5E9"><strong>待产数量</strong></td>
	<td nowrap rowspan="2" bgcolor="#8DB5E9"><strong>颜色</strong></td>
	<td nowrap rowspan="2" bgcolor="#8DB5E9"><strong>业务员</strong></td>
	<td nowrap rowspan="2" bgcolor="#8DB5E9"><strong>客户交期</strong></td>
	<td nowrap rowspan="2" bgcolor="#8DB5E9"><strong>生管交期</strong></td>
	<td nowrap rowspan="2" bgcolor="#8DB5E9"><strong>每月产能</strong></td>
	<td nowrap rowspan="2" bgcolor="#8DB5E9"><strong>过期</strong></td>
	<td nowrap colspan="5" bgcolor="#8DB5E9" height="12" align="center"><strong>订单交期分布/周次</strong></td>
	<td nowrap rowspan="2" bgcolor="#8DB5E9"><strong><%=(Month(dateadd("d",1,lastday)))%>月</strong></td>
	<td nowrap rowspan="2" bgcolor="#8DB5E9"><strong><%=(Month(dateadd("d",1,lastday))+1)%>月</strong></td>
	<td nowrap rowspan="2" bgcolor="#8DB5E9"><strong>将来</strong></td>
	<td nowrap rowspan="2" bgcolor="#8DB5E9"><strong>备注</strong></td>
</tr>
<tr height="12" align="center">
<%
	for RowsNum=GetWeekNo(now()) to (GetWeekNo(now())+4)
%>
	<td nowrap bgcolor="#8DB5E9" height="12"><strong>WK<%=RowsNum%></strong></td>
<%
	next
%>
	</tr>
<%
	Server.ScriptTimeout = 999999
	sql="select b.FBillNo as OrderID,a.FBillno as ICMOID,d.FName as Product,d.FNumber as ProductID,f.FName as productType,a.FQty,a.FQty-a.FauxStockQty as needQty,e.FName,i.fname as department,h.FDate11,c.FDate,h.FText7,g.ProductCycle,g.MonthCapacity,dateadd(d,-g.ProductCycle,c.FDate) as NeedStart,FPlanCommitDate,DATENAME(week,c.FDate) as NeedWeek "&_
"from icmo a,SEOrder b,SEOrderEntry c,t_ICItemCore d,t_Emp e,t_item f,t_dhtzdEntry h, "&_
" "&AllOPENROWSET&" zxpt.dbo.parametersys_PMProductCycle) g,t_department i "&_
"where a.FItemid=c.Fitemid and a.FSourceEntryID=c.FEntryID and a.FOrderInterID=c.FINterID "&_
"and c.FInterID=b.FinterID and a.Fitemid=d.FItemid and a.FHeadSelfJ0178=e.FItemid and a.fworkshop=i.FItemID "&_
"and g.ProductTypeId = f.FItemid and c.FSourceEntryID=h.FENtryID and h.FBase=a.Fitemid "&_
"and (left(d.FNumber,4)=f.FNumber or left(d.FNumber,7)=f.FNumber) and f.FItemClassID=4 and f.FDetail=0 "&_
"and a.fstatus<>3 and a.FCancellation=0 and a.FQty>a.FauxStockQty "
	if request("QueryStr")<>"" then sql=sql&" and "&request("QueryStr")
	sql=sql&"order by f.FNumber asc,c.FDate asc"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connk3,1,1
	RowsNum=0
	dim zzz:zzz=0
	while(not rs.eof)
	RowsNum=RowsNum+1
	tempStr1=""
response.Write("<tr bgcolor='#EBF2F9' align='right'>")
	response.Write("<td nowrap>"&RowsNum&"</td>")
	response.Write("<td nowrap>"&rs("OrderID")&"</td>")
	response.Write("<td nowrap>"&rs("ICMOID")&"</td>")
	response.Write("<td>"&rs("Product")&"</td>")
	response.Write("<td>"&rs("ProductID")&"</td>")
	response.Write("<td nowrap>"&rs("productType")&"</td>")
	response.Write("<td nowrap>"&rs("department")&"</td>")
	response.Write("<td nowrap>"&rs("FQty")&"</td>")
	response.Write("<td nowrap>"&rs("needQty")&"</td>")
	response.Write("<td nowrap>"&rs("FText7")&"</td>")
	response.Write("<td nowrap>"&rs("FName")&"</td>")
	response.Write("<td nowrap>"&rs("FDate11")&"</td>")
	response.Write("<td nowrap>"&rs("FDate")&"</td>")
	response.Write("<td nowrap>"&rs("MonthCapacity")&"</td>")
	if datediff("d",rs("FDate"),now()) >0 then
		response.Write("<td>"&rs("needQty")&"</td>")
		TotalArr(0)=TotalArr(0)+CDBl(rs("needQty"))
	else
		response.Write("<td></td>")
	end if
	for zzz=0 to 4
		if cint(rs("NeedWeek"))-GetWeekNo(now())=zzz then
		TotalArr(zzz+1)=TotalArr(zzz+1)+CDBl(rs("needQty"))
			response.Write("<td>"&rs("needQty")&"</td>")
		else
			response.Write("<td></td>")
		end if
	next
	if datediff("d",rs("FDate"),lastday) <0 then
		if Month(dateadd("d",1,lastday))=Month(rs("FDate")) then
			response.Write("<td>"&rs("needQty")&"</td>")
			response.Write("<td></td>")
			response.Write("<td></td>")
			TotalArr(zzz+1)=TotalArr(zzz+1)+CDBl(rs("needQty"))
		elseif Month(dateadd("d",1,lastday))+1=Month(rs("FDate")) then
			response.Write("<td></td>")
			response.Write("<td>"&rs("needQty")&"</td>")
			response.Write("<td></td>")
			TotalArr(zzz+2)=TotalArr(zzz+2)+CDBl(rs("FQty"))
		else		
			response.Write("<td></td>")
			response.Write("<td></td>")
			response.Write("<td>"&rs("needQty")&"</td>")
			TotalArr(zzz+3)=TotalArr(zzz+3)+CDBl(rs("needQty"))
		end if
	else
		response.Write("<td></td>")
		response.Write("<td></td>")
		response.Write("<td></td>")
	end if
	response.Write("<td></td>")
	response.Write("</tr>")
		rs.movenext
	wend
	response.Write("<tr bgcolor='#EBF2F9' align='right'>")
	response.Write("<td colspan='2'>合计</td>")
	response.Write("<td colspan='12'></td>")
	for zzz=0 to 8
		response.Write("<td>"&TotalArr(zzz)&"</td>")
	next
	response.Write("<td></td>")
	response.Write("</tr>")
%>
</table>
</div>
	<%
	rs.close
	set rs=nothing 
elseif showType="TypeSum" then 
	dim TotalArr2(9)
%>
<div id="listtable" style="width:100%; height:100%; overflow:scroll">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
<tr>
	<td class="tablemenu" colspan="35" height="20" width="100%" id="formove"><font color="#15428B"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:$('#listDiv').hide();$('#QueryTable').show();" >&nbsp;<strong>生管订单交期分布表</strong></font></td>
</tr>
<tr height="12" align="center">
	<td nowrap bgcolor="#8DB5E9"><strong>序号</strong></td>
	<td nowrap bgcolor="#8DB5E9"><strong>产品分类</strong></td>
	<td nowrap bgcolor="#8DB5E9"><strong>生产部门</strong></td>
	<td nowrap bgcolor="#8DB5E9"><strong>订单总计</strong></td>
	<td nowrap bgcolor="#8DB5E9"><strong>待产总计</strong></td>
	<td nowrap bgcolor="#8DB5E9"><strong>生产周期</strong></td>
	<td nowrap bgcolor="#8DB5E9"><strong>月产能</strong></td>
	<td nowrap bgcolor="#8DB5E9"><strong>过期</strong></td>
	<td nowrap bgcolor="#8DB5E9"><strong><%=Month(now())%>月</strong></td>
	<td nowrap bgcolor="#8DB5E9"><strong>产能结余</strong></td>
	<td nowrap bgcolor="#8DB5E9"><strong><%=Month(dateadd("m",1,now()))%>月</strong></td>
	<td nowrap bgcolor="#8DB5E9"><strong>产能结余</strong></td>
	<td nowrap bgcolor="#8DB5E9"><strong><%=Month(dateadd("m",2,now()))%>月</strong></td>
	<td nowrap bgcolor="#8DB5E9"><strong>产能结余</strong></td>
	<td nowrap bgcolor="#8DB5E9"><strong>将来</strong></td>
	<td nowrap bgcolor="#8DB5E9"><strong>备注</strong></td>
</tr>

<%
	Server.ScriptTimeout = 999999
	sql="select productType,ProductCycle,department,MonthCapacity,sum(FQty) as b0,sum(a1) as b1,sum(a2) as b2,sum(a3) as b3,sum(a4) as b4,sum(a5) as b5,(MonthCapacity-day(getdate())*MonthCapacity/30-sum(a2)) as b6,MonthCapacity-sum(a3) as b7,MonthCapacity-sum(a4) as b8,sum(needQty) as b9 from "&_
"	(select f.FName as productType,case WHEN left(i.Fnumber,2)='10' THEN '一分厂' "&_
"WHEN left(i.Fnumber,2)='11' THEN '二分厂' "&_
"WHEN left(i.Fnumber,2)='12' THEN '三分厂' "&_
"WHEN left(i.Fnumber,2)='19' or left(i.Fnumber,2)='18' THEN '娄桥分厂' "&_
"else i.FName end as department, "&_
"a.FQty,a.FQty-a.FauxStockQty as needQty,c.FDate,g.ProductCycle,g.MonthCapacity, "&_
"case when datediff(d,c.FDate,getdate())>0 then a.FQty-a.FauxStockQty else 0 end as a1, "&_
"case when datediff(d,c.FDate,getdate())<1 and datediff(mm,c.FDate,getdate())=0 then a.FQty-a.FauxStockQty else 0 end as a2, "&_
"case when datediff(mm,c.FDate,getdate())=-1 then a.FQty-a.FauxStockQty else 0 end as a3, "&_
"case when datediff(mm,c.FDate,getdate())=-2 then a.FQty-a.FauxStockQty else 0 end as a4, "&_
"case when datediff(mm,c.FDate,getdate())<-2 then a.FQty-a.FauxStockQty else 0 end as a5 "&_
"from icmo a,SEOrder b,SEOrderEntry c,t_ICItemCore d,t_Emp e,t_item f, "&_
" "&AllOPENROWSET&" zxpt.dbo.parametersys_PMProductCycle) g,t_department i "&_
"where a.FItemid=c.Fitemid and a.FSourceEntryID=c.FEntryID and a.FOrderInterID=c.FINterID "&_
"and c.FInterID=b.FinterID and a.Fitemid=d.FItemid and a.FHeadSelfJ0178=e.FItemid "&_
"and g.ProductTypeId = f.FItemid and a.fworkshop=i.FItemID "&_
"and (left(d.FNumber,4)=f.FNumber or left(d.FNumber,7)=f.FNumber) and f.FItemClassID=4 and f.FDetail=0 "&_
"and a.fstatus<>3 and a.FCancellation=0 and a.FQty>a.FauxStockQty  "
	if request("QueryStr")<>"" then sql=sql&" and "&request("QueryStr")
	sql=sql&" ) zzz group by productType,ProductCycle,MonthCapacity,department order by department,productType"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connk3,1,1
	RowsNum=0
	while(not rs.eof)
	RowsNum=RowsNum+1
	tempStr1=""
response.Write("<tr bgcolor='#EBF2F9' align='right'>")
	response.Write("<td nowrap>"&RowsNum&"</td>")
	response.Write("<td nowrap>"&rs("productType")&"</td>")
	response.Write("<td nowrap>"&rs("department")&"</td>")
	response.Write("<td nowrap>"&formatnumber(rs("b0"),0)&"</td>")
	response.Write("<td nowrap>"&formatnumber(rs("b9"),0)&"</td>")
	response.Write("<td nowrap>"&rs("ProductCycle")&"</td>")
	response.Write("<td nowrap>"&rs("MonthCapacity")&"</td>")
	response.Write("<td nowrap>"&formatnumber(rs("b1"),0)&"</td>")
	response.Write("<td nowrap>"&formatnumber(rs("b2"),0)&"</td>")
	response.Write("<td nowrap>"&formatnumber(rs("b6"),0)&"</td>")
	response.Write("<td nowrap>"&formatnumber(rs("b3"),0)&"</td>")
	response.Write("<td nowrap>"&formatnumber(rs("b7"),0)&"</td>")
	response.Write("<td nowrap>"&formatnumber(rs("b4"),0)&"</td>")
	response.Write("<td nowrap>"&formatnumber(rs("b8"),0)&"</td>")
	response.Write("<td nowrap>"&formatnumber(rs("b5"),0)&"</td>")
	response.Write("<td></td>")
	response.Write("</tr>")
	TotalArr2(0)=TotalArr2(0)+CDBl(rs("b0"))
	TotalArr2(1)=TotalArr2(1)+CDBl(rs("b1"))
	TotalArr2(2)=TotalArr2(2)+CDBl(rs("b2"))
	TotalArr2(3)=TotalArr2(3)+CDBl(rs("b3"))
	TotalArr2(4)=TotalArr2(4)+CDBl(rs("b4"))
	TotalArr2(5)=TotalArr2(5)+CDBl(rs("b5"))
	TotalArr2(6)=TotalArr2(6)+CDBl(rs("b6"))
	TotalArr2(7)=TotalArr2(7)+CDBl(rs("b7"))
	TotalArr2(8)=TotalArr2(8)+CDBl(rs("b8"))
	TotalArr2(9)=TotalArr2(9)+CDBl(rs("b9"))
		rs.movenext
	wend
	response.Write("<tr bgcolor='#EBF2F9' align='right'>")
	response.Write("<td colspan='3'>合计</td>")
	response.Write("<td>"&formatnumber(TotalArr2(0),0)&"</td>")
	response.Write("<td>"&formatnumber(TotalArr2(9),0)&"</td>")
	response.Write("<td></td>")
	response.Write("<td></td>")
	response.Write("<td>"&formatnumber(TotalArr2(1),0)&"</td>")
	response.Write("<td>"&formatnumber(TotalArr2(2),0)&"</td>")
	response.Write("<td>"&formatnumber(TotalArr2(6),0)&"</td>")
	response.Write("<td>"&formatnumber(TotalArr2(3),0)&"</td>")
	response.Write("<td>"&formatnumber(TotalArr2(7),0)&"</td>")
	response.Write("<td>"&formatnumber(TotalArr2(4),0)&"</td>")
	response.Write("<td>"&formatnumber(TotalArr2(8),0)&"</td>")
	response.Write("<td>"&formatnumber(TotalArr2(5),0)&"</td>")
	response.Write("<td></td>")
	response.Write("</tr>")
%>
</table>
</div>
	<%
	rs.close
	set rs=nothing 
elseif showType="getInfo" then 
	sql="select b.FName from parametersys_PMProductCycle a, "&AllOPENROWSET&" AIS20081217153921.dbo.t_item) as b "&_
	"where b.FItemClassId=4 and b.FDetail=0 and a.ProductTypeId=b.FItemid"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
	response.Write("[")
	do until rs.eof
		response.Write("{""CValue"":"""&rs("FName")&""",""title"":"""&rs("FName")&"""}")
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
