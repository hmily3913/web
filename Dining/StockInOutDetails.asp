<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurviewFLW"),"|603,")=0 then 
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
      datafrom=" Dining_StockInOut a,Dining_Material b "
  dim datawhere'数据条件
  dim i'用于循环的整数
    datawhere=" where a.MaterialId=b.SerialNum "
	Dim searchterm,searchcols
	
	if Request.Form("query") <> "" then
	searchterm = Request.Form("query")
	searchcols = Request.Form("qtype")
	if isnumeric(searchterm) then
	datawhere = datawhere&" and " & searchcols & " = " & searchterm & " "
	else
	datawhere = datawhere&" and " & searchcols & " LIKE '%" & searchterm & "%' "
	end if
	End if
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
	Dim sortname
	if Request.Form("sortname") = "" then
	sortname = "a.SerialNum" 
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
    sql="select a.SerialNum from "& datafrom &" " & datawhere & taxis
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
    sql="select a.*,b.MaterialName,b.Unit,b.StockQty from "& datafrom &" where a.MaterialId=b.SerialNum and a.SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    do until rs.eof'填充数据到表格'
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("InOutDate")%>","<%=rs("InOutType")%>","<%=rs("MaterialId")%>","<%=rs("MaterialName")%>","<%=rs("Unit")%>","<%=rs("StockQty")%>","<%=rs("InOutPrice")%>","<%=rs("InOutQty")%>","<%=rs("CheckFlag")%>","<%=rs("Remark")%>","<%=rs("Biller")%>"]}
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
'-----------------------------------------------------------'
elseif showType="DataProcess" then 
  detailType=request("detailType")
  if detailType="AddNew" then
  	if  Instr(session("AdminPurviewFLW"),"|603.1,")>0 then
			for   i=2   to   Request.form("SerialNum").count
				if (Request.form("DeleteFlag")(i)="0" or Request.form("DeleteFlag")(i)="") and Request.form("MaterialId")(i)<>"" then
					connzxpt.Execute("insert into Dining_StockInOut (InOutDate,MaterialId,InOutPrice,InOutQty,InOutType,Remark,BillerID,Biller,BillDate) values ('"&Request.form("InOutDate")&"','"&Request.form("MaterialId")(i)&"','"&Request.form("InOutPrice")(i)&"','"&Request.form("InOutQty")(i)&"','"&Request.form("InOutType")&"','"&Request.form("Remark")(i)&"','"&UserName&"','"&AdminName&"','"&now()&"')")
				end if
			next
			response.write "###"
		else
			response.write ("你没有权限进行此操作！")
			response.end
		end if
  elseif detailType="Edit"  then
		if Instr(session("AdminPurviewFLW"),"|603.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		for   i=2   to   Request.form("SerialNum").count
			if Request.form("DeleteFlag")(i)="1" and Request.Form("SerialNum")(i)<>"" then
				connzxpt.Execute("Delete from Dining_StockInOut where SerialNum="&Request.Form("SerialNum")(i))
			elseif Request.Form("SerialNum")(i)<>"" then
				connzxpt.Execute("update Dining_StockInOut set InOutType='"&Request.form("InOutType")&"',InOutDate='"&Request.form("InOutDate")&"',MaterialId='"&Request.form("MaterialId")(i)&"',InOutPrice='"&Request.form("InOutPrice")(i)&"',InOutQty='"&Request.form("InOutQty")(i)&"',Remark='"&Request.form("Remark")(i)&"',BillerID='"&UserName&"',Biller='"&AdminName&"',BillDate='"&now()&"' where SerialNum="&Request.Form("SerialNum")(i))
			else
					connzxpt.Execute("insert into Dining_StockInOut (InOutDate,MaterialId,InOutPrice,InOutQty,InOutType,Remark,BillerID,Biller,BillDate) values ('"&Request.form("InOutDate")&"','"&Request.form("MaterialId")(i)&"','"&Request.form("InOutPrice")(i)&"','"&Request.form("InOutQty")(i)&"','"&Request.form("InOutType")&"','"&Request.form("Remark")(i)&"','"&UserName&"','"&AdminName&"','"&now()&"')")
			end if
		next
		response.write "###"
  elseif detailType="Delete" then
		if Instr(session("AdminPurviewFLW"),"|603.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
    SerialNum=request("SerialNum")
		sql="select * from Dining_StockInOut where CheckFlag=1 and SerialNum in ("&SerialNum&")"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		if rs.eof then
			connzxpt.Execute("Delete from Dining_StockInOut where SerialNum in ("&SerialNum&")")
			response.write "###"
		else
			response.Write("已审核不允许删除！")
			response.End()
		end if
  elseif detailType="Check"  then
		SerialNum=request("SerialNum")
		if Instr(session("AdminPurviewFLW"),"|603.2,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		sql="select * from Dining_StockInOut where SerialNum in ("&SerialNum&")"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
			if rs("CheckFlag")=0 then
				rs("Checker")=AdminName
				rs("CheckerID")=UserName
				rs("CheckDate")=now()
				rs("CheckFlag")=1
				if rs("InOutType")="售出" then
					connzxpt.Execute("update Dining_Material set StockQty=StockQty+"&rs("InOutQty")&",UseFlag=1 where SerialNum ="&rs("MaterialId"))
				elseif rs("InOutType")="购入" then
					connzxpt.Execute("update Dining_Material set Price="&rs("InOutPrice")&",CostPrice=(CostPrice*StockQty+"&(Cdbl(rs("InOutPrice"))*Cdbl(rs("InOutQty")))&")/(StockQty+"&rs("InOutQty")&"),StockQty=StockQty+"&rs("InOutQty")&",UseFlag=1 where SerialNum ="&rs("MaterialId"))
				end if
			end if		
			rs.movenext
		wend
		rs.update
		rs.close
		set rs=nothing 
		response.write "###"
  elseif detailType="unCheck"  then
		if Instr(session("AdminPurviewFLW"),"|603.2,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		SerialNum=request("SerialNum")
		connzxpt.Execute("update Dining_StockInOut set CheckFlag=0 where SerialNum in ("&SerialNum&")")
		response.write "###"
  end if
elseif showType="getInfo" then 
	sql="select SerialNum,MaterialName,Unit,Price,StockQty from Dining_Material where Type<>'菜品' and ForbidFlag=0 "
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
	response.write "["
	do until rs.eof
	Response.Write("{")
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
	Response.Write("]")
	rs.close
	set rs=nothing 
elseif showType="Excel" then 
	sql="exec sp_Dininginout "&request("Year")&","&request("Month")&" "
	response.Write(sql)
	if request("Printtag")=1 then
	response.ContentType("application/vnd.ms-excel")
	response.AddHeader "Content-disposition", "attachment; filename=erpData.xls"
	end if
%>
<div id="listtable" style="width:100%; height:420; overflow:scroll">
<table>
<tbody id="TbDetails">
<%
  set rs=server.createobject("adodb.recordset")
  rs.open sql,Connzxpt,0,1
	if not rs.eof then
%>
<tr>    <td height="20" width="100%" class="tablemenu" colspan="<%=rs.fields.count%>"><font color="#15428B"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="$('#listtable').hide().css('z-index','550');$('#QueryTable').show();" >&nbsp;<strong>页面查看明细</strong></font></td>
</tr>
<tr bgcolor="#99BBE8">
<%
	  for i=0 to rs.fields.count-1
		response.write ("<td nowrap=""nowrap"">"&rs.fields(i).name&"</td>")
	  next
		response.Write("</tr>" & vbCrLf)
		rs.movenext
	end if
	while(not rs.eof)
		response.Write("<tr bgcolor=""#EBF2F9"">")
	  for i=0 to rs.fields.count-1
	    if IsNull(rs.fields(i).value) then
		response.write ("<td>"&rs.fields(i).value&"</td>")
		elseif rs.fields(i).value="0" then
		response.write ("<td></td>")
		else
		response.write ("<td>"&JsonStr(rs.fields(i).value)&"</td>")
		end if
	  next
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
