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
	sql="select count(1) as idcount from Calendar a where EXISTS (select * from Calendar b where a.weeknum-b.weeknum<5 and a.weeknum-b.weeknum>-1 and datediff(d,b.date,getdate())=0 and datediff(yy,a.date,getdate())=0 and datediff(d,a.date,getdate())<1)"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
	daycount=rs("idcount")
	dim TotalArr(35)
%>
<div id="listtable" style="width:100%; height:100%; overflow:scroll">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
<tr>
	<td class="tablemenu" colspan="35" height="20" width="100%" id="formove"><font color="#15428B"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:$('#listDiv').hide();$('#QueryTable').show();" >&nbsp;<strong>生管投产计划表</strong></font></td>
</tr>
<tr height="12" align="center">
	<td nowrap rowspan="3" bgcolor="#8DB5E9"><strong>序号</strong></td>
	<td nowrap rowspan="3" bgcolor="#8DB5E9"><strong>销售订单号</strong></td>
	<td nowrap rowspan="3" bgcolor="#8DB5E9"><strong>生产任务单</strong></td>
	<td nowrap rowspan="3" bgcolor="#8DB5E9"><strong>产品名称</strong></td>
	<td nowrap rowspan="3" bgcolor="#8DB5E9"><strong>产品分类</strong></td>
	<td nowrap rowspan="3" bgcolor="#8DB5E9"><strong>生产部门</strong></td>
	<td nowrap rowspan="3" bgcolor="#8DB5E9"><strong>订单数量</strong></td>
	<td nowrap rowspan="3" bgcolor="#8DB5E9"><strong>待产数量</strong></td>
	<td nowrap rowspan="3" bgcolor="#8DB5E9"><strong>颜色</strong></td>
	<td nowrap rowspan="3" bgcolor="#8DB5E9"><strong>业务员</strong></td>
	<td nowrap rowspan="3" bgcolor="#8DB5E9"><strong>客户交期</strong></td>
	<td nowrap rowspan="3" bgcolor="#8DB5E9"><strong>生管交期</strong></td>
	<td nowrap rowspan="3" bgcolor="#8DB5E9"><strong>生产周期</strong></td>
	<td nowrap rowspan="3" bgcolor="#8DB5E9"><strong>过期任务单</strong></td>
	<td nowrap colspan="<%=daycount%>" bgcolor="#8DB5E9" height="12" align="center"><strong>投产计划/周次</strong></td>
	<td nowrap rowspan="3" bgcolor="#8DB5E9"><strong>将来</strong></td>
	<td nowrap rowspan="3" bgcolor="#8DB5E9"><strong>备注</strong></td>
</tr>
<tr height="12" align="center">
<%
	
	sql="select date,weeknum from Calendar a where EXISTS (select * from Calendar b where a.weeknum-b.weeknum<5 and a.weeknum-b.weeknum>-1 and datediff(d,b.date,getdate())=0 and datediff(yy,a.date,getdate())=0 and datediff(d,a.date,getdate())<1)"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
	while(not rs.eof)
		RowsNum=RowsNum+1
		if rs("weeknum")<>tempStr1 then
			if tempStr1<>"临时字符串" then
			
%>
	<td nowrap colspan="<%=RowsNum%>" bgcolor="#8DB5E9" height="12"><strong>WK<%=(rs("weeknum")-1)%></strong></td>
<%
				RowsNum=0
			end if
			tempStr1=rs("weeknum")
		end if
		lastweek=rs("weeknum")
		rs.movenext
	wend
	rs.movefirst
%>
	<td nowrap colspan="7" bgcolor="#8DB5E9" height="12"><strong>WK<%=lastweek%></strong></td>
	</tr>
	<tr height="12" align="center">
<%
	 while(not rs.eof)
%>
	<td bgcolor="#8DB5E9" height="12"><strong><%=MONTH(rs("date"))&"-"&day(rs("date"))%></strong></td>
<%
		lastday=rs("date")
		rs.movenext
	wend
%>
	</tr>
<%
	sql="select b.FBillNo as OrderID,a.FBillno as ICMOID,d.FName as Product,f.FName as productType,a.FQty,a.FQty-a.FauxStockQty as needQty,e.FName,h.FDate11,c.FDate,h.FText7,g.ProductCycle,g.MonthCapacity,dateadd(d,-g.ProductCycle,c.FDate) as NeedStart,FPlanCommitDate,i.FName as department "&_
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
	response.Write("<td nowrap>"&rs("productType")&"</td>")
	response.Write("<td nowrap>"&rs("department")&"</td>")
	response.Write("<td nowrap>"&rs("FQty")&"</td>")
	response.Write("<td nowrap>"&rs("needQty")&"</td>")
	response.Write("<td nowrap>"&rs("FText7")&"</td>")
	response.Write("<td nowrap>"&rs("FName")&"</td>")
	response.Write("<td nowrap>"&rs("FDate11")&"</td>")
	response.Write("<td nowrap>"&rs("FDate")&"</td>")
	response.Write("<td nowrap>"&rs("ProductCycle")&"</td>")
	tempStr1=""
	if datediff("d",rs("FPlanCommitDate"),now()) >0 then tempStr1=" title='当前计划日期"&rs("FPlanCommitDate")&"'  bgcolor='#FF0000'"
	if datediff("d",rs("NeedStart"),now()) >0 then
		response.Write("<td "&tempStr1&">"&rs("needQty")&"</td>")
		TotalArr(0)=TotalArr(0)+CDBl(rs("needQty"))
	else
		response.Write("<td "&tempStr1&"></td>")
	end if
	for zzz=0 to daycount-1
		tempStr1=""
		if datediff("d",now(),rs("FPlanCommitDate"))=zzz then tempStr1=" title='当前计划日期"&rs("FPlanCommitDate")&"'  bgcolor='#FF0000'"
		if datediff("d",now(),rs("NeedStart"))=zzz then
		TotalArr(zzz+1)=TotalArr(zzz+1)+CDBl(rs("needQty"))
			response.Write("<td "&tempStr1&">"&rs("needQty")&"</td>")
		else
			response.Write("<td "&tempStr1&"></td>")
		end if
	next
	tempStr1=""
	if datediff("d",rs("FPlanCommitDate"),lastday) <0 then tempStr1=" title='当前计划日期"&rs("FPlanCommitDate")&"'  bgcolor='#FF0000'"
	if datediff("d",rs("NeedStart"),lastday) <0 then
		response.Write("<td "&tempStr1&">"&rs("needQty")&"</td>")
		TotalArr(zzz+1)=TotalArr(zzz+1)+CDBl(rs("needQty"))
	else
		response.Write("<td "&tempStr1&"></td>")
	end if
	response.Write("<td></td>")
	response.Write("</tr>")
		rs.movenext
	wend
	response.Write("<tr bgcolor='#EBF2F9' align='right'>")
	response.Write("<td colspan='2'>合计</td>")
	response.Write("<td colspan='11'></td>")
	for zzz=0 to daycount+1
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
end if
 %>
