<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
if Instr(session("AdminPurview"),"|106,")=0 then 
  response.write "{""page"":""1"",""total"":""1"",""rows"":[{""id"":""0"",""cell"":["
  response.write ("""<font color='red')>你不具有该管理模块的操作权限，请返回！</font>""]}]}")
  response.end
end if
'========判断是否具有管理权限
dim showType,start_date,end_date,print_tag,UserName,AdminName,Depart
UserName=session("UserName")
AdminName=session("AdminName")
showType=request("showType")
print_tag=request("print_tag")
Depart=session("Depart")
if print_tag=1 then
response.ContentType("application/vnd.ms-excel")
response.AddHeader "Content-disposition", "attachment; filename=erpData.xls"
end if
if showType="DetailsList" then 
  dim page'页码
      page=clng(request("page"))
  dim idCount'记录总数
  dim pages'每页条数
      pages=request("rp")
  dim pagec'总页数

  dim datafrom'数据表名
      datafrom=" sale_OrderForecast "
  dim datawhere'数据条件
    datawhere=" where 1=1 "
	Dim searchterm,searchcols
	if Request.Form("query") <> "" then
	searchterm = Request.Form("query")
	searchcols = request.Form("qtype")
		if Request.Form("qtype")="CtrlDate" then
			datawhere = datawhere&" and datediff(d,"&searchcols&",'"&searchterm&"')=0 "
		else
			datawhere = datawhere&" and "&searchcols&" like '%"&searchterm&"%' "
		end if
	End if
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
	Dim sortname
	if Request.Form("sortname") = "" then
	sortname = "SerialNum" 
	Else
	sortname = Request.Form("sortname")
	End If
	Dim sortorder
	if Request.Form("sortorder") = "" then
	sortorder = "desc"
	Else
	sortorder = Request.Form("sortorder")
	End If
      taxis=" order by "&sortname&" "&sortorder
  dim i'用于循环的整数
  dim rs,sql'sql语句
  '获取记录总数
  sql="select count(1) as idCount from "& datafrom &" " & datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,0,1
  idCount=rs("idCount")
  if(idcount>0) then'如果记录总数=0,则不处理
    if(idcount mod pages=0)then'如果记录总数除以每页条数有余数,则=记录总数/每页条数+1
	  pagec=int(idcount/pages)'获取总页数
   	else
      pagec=int(idcount/pages)+1'获取总页数
    end if
  end if
	'获取本页需要用到的id============================================
    '读取所有记录的id数值,因为只有id所以速度很快
    sql="select SerialNum from "& datafrom &" " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    rs.pagesize = pages '每页显示记录数
	rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
	  if(i=1)then
	    sqlid=rs("SerialNum")
	  else
	    sqlid=sqlid &","&rs("SerialNum")
	  end if
	  rs.movenext
    next
  '获取本页需要用到的id结束============================================
'-----------------------------------------------------------
'-----------------------------------------------------------
%>
{"page":"<%=page%>","total":"<%=idcount%>","rows":[
<%
  if sqlid<>"" then'如果记录总数=0,则不处理
    '用in刷选本页所语言的数据,仅读取本页所需的数据,所以速度快
    sql="select * from "& datafrom &" where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    do until rs.eof'填充数据到表格
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("Register")%>","<%=rs("RegisterName")%>","<%=rs("Departmentname")%>","<%=rs("YearNum")%>","<%=rs("WeekNum")%>","<%=rs("AllQty")%>","<%=rs("Biller")%>","<%=rs("BillDate")%>","<%=rs("Checker")%>","<%=rs("CheckDate")%>","<%=rs("CheckFlag")%>"]}
<%		
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
  end if
  rs.close
  set rs=nothing
response.Write"]}"
'-----------------------------------------------------------
elseif showType="DataProcess" then 
  detailType=request("detailType")
  if detailType="AddNew" then
  	if  Instr(session("AdminPurview"),"|106.1,")>0 then
			set rs = server.createobject("adodb.recordset")
			sql="select * from sale_OrderForecast"
			rs.open sql,connzxpt,1,3
			rs.addnew
			rs("Biller")=AdminName
			rs("BillerID")=UserName
			rs("BillDate")=now()
			rs("Register")=Request("Register")
			rs("RegisterName")=Request("RegisterName")
			rs("Department")=Request("Department")
			rs("Departmentname")=Request("Departmentname")
			rs("YearNum")=Request("YearNum")
			rs("WeekNum")=Request("WeekNum")
			rs.update
			set rs=connzxpt.Execute("select top 1 SerialNum from sale_OrderForecast order by serialnum desc")
			if rs("SerialNum")="" then
			SerialNum=10000000
			else
			SerialNum=rs("SerialNum")
			end if
			dim allqty:allqty=0
			for   i=2   to   Request.form("SerialNumD").count
				connzxpt.Execute("insert into sale_OrderForecastDetails (SNum,ProductTypeId,Forecasts,Remark) values ('"&SerialNum&"','"&Request.form("ProductTypeId")(i)&"','"&Request.form("Forecasts")(i)&"','"&Request.form("Remark")(i)&"')")
				allqty=allqty+cdbl(Request.form("Forecasts")(i))
			next
			connzxpt.Execute("update sale_OrderForecast set AllQty="&allqty&" where serialnum="&SerialNum)
			rs.close
			set rs=nothing 
			response.write "保存成功！"
		else
			response.write ("你没有权限进行此操作！")
			response.end
		end if
  elseif detailType="Edit"  then
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select * from sale_OrderForecast where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurview"),"|106.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
			end if
		if rs("BillerID")<>UserName and rs("Register")<>UserName then
			response.write ("只能编辑自己添加的数据！")
			response.end
		end if
		if rs("CheckFlag")>0 then
			response.write ("当前状态不允许编辑，请检查！")
			response.end
		end if
		rs("Biller")=AdminName
		rs("BillerID")=UserName
		rs("BillDate")=now()
		rs("Register")=Request("Register")
		rs("RegisterName")=Request("RegisterName")
		rs("Department")=Request("Department")
		rs("Departmentname")=Request("Departmentname")
		rs("YearNum")=Request("YearNum")
		rs("WeekNum")=Request("WeekNum")
    rs.update
		allqty=0
		for   i=2   to   Request.form("SerialNumD").count
				connzxpt.Execute("update sale_OrderForecastDetails set Forecasts='"&Request.form("Forecasts")(i)&"',Remark='"&Request.form("Remark")(i)&"' where SerialNumD="&Request.Form("SerialNumD")(i))
				allqty=allqty+cdbl(Request.form("Forecasts")(i))
		next
		connzxpt.Execute("update sale_OrderForecast set AllQty="&allqty&" where serialnum="&SerialNum)
		rs.close
		set rs=nothing 
		response.write "修改成功！"
  elseif detailType="Delete" then
		if Instr(session("AdminPurview"),"|106.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
		sql="select * from sale_OrderForecast where CheckFlag>0 and SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,1
		if rs.bof and rs.eof then
			connzxpt.Execute("Delete from sale_OrderForecast where SerialNum in ("&SerialNum&")")
			connzxpt.Execute("Delete from sale_OrderForecastDetails where SNum in ("&SerialNum&")")
			response.write "删除成功！"
		else
			response.write ("已经审核不允许删除！")
			response.end
		end if
		rs.close
		set rs=nothing 
  elseif detailType="Check" then
		if Instr(session("AdminPurview"),"|106.2,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select * from sale_OrderForecast where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
			rs("Checker")=AdminName
			rs("CheckerID")=UserName
			rs("CheckDate")=now()
			rs("CheckFlag")=1
			rs.movenext
		wend
		rs.update
		rs.close
		set rs=nothing 
		response.write "审核成功！"
  elseif detailType="unCheck" then
		if Instr(session("AdminPurview"),"|106.2,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select * from sale_OrderForecast where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
			if rs("CheckFlag")=0 then
				response.write ("此单据未审核，不允许反审核！")
				response.end
			end if
			rs("Checker")=AdminName
			rs("CheckerID")=UserName
			rs("CheckDate")=now()
			rs("CheckFlag")=0
			rs.movenext
		wend
		rs.update
		rs.close
		set rs=nothing 
		response.write "驳回成功！"
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="Register" then
    InfoID=request("InfoID")
		sql="select top 1 a.*,b.部门名称 from [N-基本资料单头] a,[G-部门资料表] b where a.二级部门=b.部门代号 and a.员工代号 like '%"&InfoID&"%' "
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
				response.write ("员工编号不存在！")
				response.end
		else
			response.write(rs("员工代号")&"###"&rs("姓名")&"###"&rs("二级部门")&"###"&rs("部门名称")&"###"&rs("性别"))
		end if
		rs.close
		set rs=nothing 
  elseif detailType="SerialNum" then
    InfoID=request("InfoID")
		sql="select * from sale_OrderForecast where SerialNum="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		if rs.bof and rs.eof then
			response.write ("对应单据不存在，请检查！")
			response.end
		else
			response.write "{""Info"":""###"",""fieldValue"":[{"
			for i=0 to rs.fields.count-1
			if IsNull(rs.fields(i).value) then
			response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
			else
			response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
			end if
			next
			response.write ("""Entrys"":[")
			sql="select a.*,b.FNumber,b.FName ,isnull(sum(FQty),0) as Qty,case when a.Forecasts=0 then 0 else round(isnull(sum(FQty),0)/a.Forecasts*100,2) end as Bili "
			sql=sql&" from sale_OrderForecastDetails a  "
			sql=sql&" inner join AIS20081217153921.dbo.t_item b on b.FItemClassId=4 and b.FDetail=0 and a.ProductTypeId=b.FItemid and SNum="&InfoID
			sql=sql&" inner join AIS20081217153921.dbo.t_item c on c.FNumber='"&rs("Register")&"'  "
			sql=sql&" left join ( "
			sql=sql&" select e.FQty,d.FBase3,f.FNumber from AIS20081217153921.dbo.t_DHTZD d , "
			sql=sql&" AIS20081217153921.dbo.t_DHTZDEntry e ,AIS20081217153921.dbo.t_icitem f "
			sql=sql&" where year(d.fdate1)="&rs("YearNum")&" and DATEPART(ww, d.fdate1)="&rs("WeekNum")&"+1 and d.FUser>0 and d.FID=e.FID "
			sql=sql&" and e.FBase=f.Fitemid ) h on h.FBase3=c.FItemid and h.FNumber like b.Fnumber+'%' "
			sql=sql&" group by a.SerialNumD,a.ProductTypeId,a.SNum,a.Remark,a.Forecasts,b.FNumber,b.FName order by b.FNumber"
			set rs=server.createobject("adodb.recordset")
			rs.open sql,connzxpt,1,1
			do until rs.eof
				response.write "{"
				for i=0 to rs.fields.count-2
				if IsNull(rs.fields(i).value) then
				response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
				else
				response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
				end if
				next
				if IsNull(rs.fields(i).value) then
				response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&"""}")
				else
				response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&"""}")
				end if
				rs.movenext
			If Not rs.eof Then
				Response.Write ","
			End If
			loop
			response.write "]}]}"
		end if
		rs.close
		set rs=nothing 
  elseif detailType="ProductType" then
		sql="select a.ProductTypeId,b.FNumber,b.FName,0 as Forecasts,'' as Remark from parametersys_PMProductCycle a,AIS20081217153921.dbo.t_item b where b.FItemClassId=4 and b.FDetail=0 and a.ProductTypeId=b.FItemid order by b.FNumber"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		response.Write("[")
    do until rs.eof'填充数据到表格
%>		
		{"ProductTypeId":"<%=rs("ProductTypeId")%>","FNumber":"<%=rs("FNumber")%>","FName":"<%=rs("FName")%>","Forecasts":"0","Remark":"","Qty":"0","Bili":"0"}
<%		
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
		rs.close
		set rs=nothing
		response.Write"]"
  end if
elseif showType="Excel" then
	sql=" select ProductID as 分类编号,ProductName as 分类名称,isnull(sum(case when DepartmentName='内贸科' then Forecsum else 0 end ),0) as 内贸预测订单, "
	sql=sql&" isnull(sum(case when DepartmentName='内贸科' then actsum else 0 end),0) as 内贸实际接单, "
	sql=sql&" isnull(sum(case when DepartmentName='外贸科' then Forecsum else 0 end),0) as 外贸预测订单, "
	sql=sql&" isnull(sum(case when DepartmentName='外贸科' then actsum else 0 end ),0) as 外贸实际接单, "
	sql=sql&" isnull(sum(Forecsum),0) 合计预测订单, "
	sql=sql&" isnull(sum(actsum),0) 合计实际接单 "
	sql=sql&" from (select g.Forecsum,g.FNumber as ProductID,g.FName as ProductName,g.Register,g.DepartmentName,g.YearNum,g.WeekNum,c.Fnumber,c.Fname,sum(h.FQty) as actsum from ( "
	sql=sql&" select sum(a.Forecasts) as Forecsum,b.FNumber,b.FName,o.Register,"&request("Year")&" as YearNum,"&request("Week")&" as WeekNum,o.DepartmentName "
	sql=sql&" from sale_OrderForecastDetails a, AIS20081217153921.dbo.t_item b,sale_OrderForecast o "
	sql=sql&" where b.FItemClassId=4 and b.FDetail=0 and a.ProductTypeId=b.FItemid and a.SNum=o.Serialnum and o.YearNum="&request("Year")&" and o.WeekNum="&request("Week")&" "
	sql=sql&" group by b.FNumber,b.FName,o.Register,o.DepartmentName) g "
	sql=sql&" inner join AIS20081217153921.dbo.t_item c on c.FNumber=g.Register "
	sql=sql&" left join (select e.FQty,d.FBase3,f.FNumber from AIS20081217153921.dbo.t_DHTZD d , "
	sql=sql&" AIS20081217153921.dbo.t_DHTZDEntry e ,AIS20081217153921.dbo.t_icitem f  "
	sql=sql&" where d.FUser>0 and d.FID=e.FID and e.FBase=f.Fitemid and year(d.fdate1)="&request("Year")&"  "
	sql=sql&" and DATEPART(ww, d.fdate1)="&request("Week")&"+1) h  "
	sql=sql&" on h.FNumber like g.Fnumber+'%' and c.Fitemid=h.FBase3 "
	sql=sql&" group by g.Forecsum,g.FNumber,g.FName,g.Register,g.DepartmentName,g.YearNum,g.WeekNum,c.Fnumber,c.Fname "
	sql=sql&" ) ccc group by ProductID,ProductName,YearNum,WeekNum order by ProductID "
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
%>
<div id="listtable" style="width:100%; height:420; overflow:scroll">
<table>
<tbody id="TbDetails">
<%
	if not rs.eof then
%>
<tr>    <td height="20" width="100%" class="tablemenu" colspan="11"><font color="#15428B"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="$('#listtable').hide().css('z-index','550');" >&nbsp;<strong><%=request("Year")%>年<%=request("Week")%>周预测汇总</strong></font></td>
</tr>
<tr bgcolor="#99BBE8">
<td nowrap="nowrap">产品类别编号</td>
<td nowrap="nowrap">产品类别名称</td>
<td nowrap="nowrap">内贸预测订单</td>
<td nowrap="nowrap">内贸实际接单</td>
<td nowrap="nowrap">内贸接单比例</td>
<td nowrap="nowrap">外贸预测订单</td>
<td nowrap="nowrap">外贸实际接单</td>
<td nowrap="nowrap">外贸接单比例</td>
<td nowrap="nowrap">合计预测订单</td>
<td nowrap="nowrap">合计实际接单</td>
<td nowrap="nowrap">合计接单比例</td>
<%
	end if
	while(not rs.eof)
		response.Write("<tr bgcolor=""#EBF2F9"">")
		response.write ("<td>"&rs.fields(0).value&"</td>")
		response.write ("<td>"&rs.fields(1).value&"</td>")
		response.write ("<td>"&rs.fields(2).value&"</td>")
		response.write ("<td>"&rs.fields(3).value&"</td>")
		if cdbl(rs.fields(2).value)=0 then
			response.write ("<td></td>")
		else
			response.write ("<td>"&formatnumber(cdbl(rs.fields(3).value)*100/cdbl(rs.fields(2).value),2)&"</td>")
		end if
		response.write ("<td>"&rs.fields(4).value&"</td>")
		response.write ("<td>"&rs.fields(5).value&"</td>")
		if cdbl(rs.fields(4).value)=0 then
			response.write ("<td></td>")
		else
			response.write ("<td>"&formatnumber(cdbl(rs.fields(5).value)*100/cdbl(rs.fields(4).value),2)&"</td>")
		end if
		response.write ("<td>"&rs.fields(6).value&"</td>")
		response.write ("<td>"&rs.fields(7).value&"</td>")
		if cdbl(rs.fields(6).value)=0 then
			response.write ("<td></td>")
		else
			response.write ("<td>"&formatnumber(cdbl(rs.fields(7).value)*100/cdbl(rs.fields(6).value),2)&"</td>")
		end if
		response.Write("</tr>" & vbCrLf)
		rs.movenext
	wend
%>
</tbody>
</table>
</div>
<%
	rs.close
	set rs=nothing
end if
 %>
