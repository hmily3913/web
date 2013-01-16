<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurviewFLW"),"|601,")=0 then 
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
      datafrom=" Dining_Material "
  dim datawhere'数据条件
  dim i'用于循环的整数
    datawhere=" where 1=1 "
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
    do until rs.eof'填充数据到表格'
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("MaterialName")%>","<%=rs("Model")%>","<%=rs("Type")%>","<%=rs("Unit")%>","<%=rs("LIFECYCLE")%>","<%=rs("Price")%>","<%=rs("CostPrice")%>","<%=rs("StockQty")%>","<%=rs("ForbidFlag")%>","<%=rs("Remark")%>","<%=rs("Biller")%>"]}
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
  	if  Instr(session("AdminPurviewFLW"),"|601.1,")>0 then
		for   i=2   to   Request.form("SerialNum").count
			if Request.form("DeleteFlag")(i)="0" and Request.form("MaterialName")(i)<>"" then
				sql="select * from Dining_Material where MaterialName='"&Request.form("MaterialName")(i)&"'"
				set rs=server.createobject("adodb.recordset")
				rs.open sql,connzxpt,1,1
				if rs.eof then
					connzxpt.Execute("insert into Dining_Material (MaterialName,Model,Type,Unit,LIFECYCLE,Price,Remark,BillerID,Biller,BillDate) values ('"&Request.form("MaterialName")(i)&"','"&Request.form("Model")(i)&"','"&Request.form("Type")(i)&"','"&Request.form("Unit")(i)&"','"&Request.form("LIFECYCLE")(i)&"','"&Request.form("Price")(i)&"','"&Request.form("Remark")(i)&"','"&UserName&"','"&AdminName&"','"&now()&"')")
				end if
				rs.close
				set rs=nothing
			end if
		next
		response.write "###"
	else
		response.write ("你没有权限进行此操作！")
		response.end
	end if
  elseif detailType="Edit"  then
		if Instr(session("AdminPurviewFLW"),"|601.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		for   i=2   to   Request.form("SerialNum").count
			if Request.form("DeleteFlag")(i)="1" and Request.Form("SerialNum")(i)<>"" then
				connzxpt.Execute("Delete from Dining_Material where SerialNum="&Request.Form("SerialNum")(i))
			elseif Request.Form("SerialNum")(i)<>"" then
				connzxpt.Execute("update Dining_Material set MaterialName='"&Request.form("MaterialName")(i)&"',Model='"&Request.form("Model")(i)&"',Type='"&Request.form("Type")(i)&"',Unit='"&Request.form("Unit")(i)&"',LIFECYCLE='"&Request.form("LIFECYCLE")(i)&"',Price='"&Request.form("Price")(i)&"',Remark='"&Request.form("Remark")(i)&"',BillerID='"&UserName&"',Biller='"&AdminName&"',BillDate='"&now()&"' where SerialNum="&Request.Form("SerialNum")(i))
			else
				sql="select * from Dining_Material where MaterialName='"&Request.form("MaterialName")(i)&"'"
				set rs=server.createobject("adodb.recordset")
				rs.open sql,connzxpt,1,1
				if rs.eof then
					connzxpt.Execute("insert into Dining_Material (MaterialName,Model,Type,Unit,LIFECYCLE,Price,Remark,BillerID,Biller,BillDate) values ('"&Request.form("MaterialName")(i)&"','"&Request.form("Model")(i)&"','"&Request.form("Type")(i)&"','"&Request.form("Unit")(i)&"','"&Request.form("LIFECYCLE")(i)&"','"&Request.form("Price")(i)&"','"&Request.form("Remark")(i)&"','"&UserName&"','"&AdminName&"','"&now()&"')")
				end if
				rs.close
				set rs=nothing
			end if
		next
		response.write "###"
  elseif detailType="Delete" then
		if Instr(session("AdminPurviewFLW"),"|601.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
    SerialNum=request("SerialNum")
		sql="select * from Dining_Material where UseFlag=1 and SerialNum in ("&SerialNum&")"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		if rs.eof then
			connzxpt.Execute("Delete from Dining_Material where SerialNum in ("&SerialNum&")")
			response.write "###"
		else
			response.Write("已使用不允许删除！")
			response.End()
		end if
		rs.close
		set rs=nothing
  elseif detailType="Forbid"  then
		if Instr(session("AdminPurviewFLW"),"|601.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		connzxpt.Execute("update Dining_Material set ForbidFlag=1 where SerialNum in ("&request("SerialNum")&")")
		response.write "###"
  elseif detailType="ForUse"  then
		if Instr(session("AdminPurviewFLW"),"|601.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		connzxpt.Execute("update Dining_Material set ForbidFlag=0 where SerialNum in ("&request("SerialNum")&")")
		response.write "###"
  end if
elseif showType="Export" then 
	set rs=server.createobject("adodb.recordset")
	sql="select * from Dining_Material where ForbidFlag=0 and Type<>'菜品'"
	rs.open sql,connzxpt,1,1
	%>
  <table border="1">
  	<tr>
    	<td>物料名称</td>
    	<td>规格</td>
    	<td>类型</td>
    	<td>单位</td>
    	<td>保质期</td>
    	<td>最新单价</td>
    	<td>平均价</td>
    	<td>库存数</td>
    	<td>备注</td>
    	<td>登记人</td>
    </tr>
	<%
	while (not rs.eof)
	%>
	  <tr>
      <td><%=rs("MaterialName")%></td>
      <td><%=rs("Model")%></td>
      <td><%=rs("Type")%></td>
      <td><%=rs("Unit")%></td>
      <td><%=rs("LIFECYCLE")%></td>
      <td><%=rs("Price")%></td>
      <td><%=rs("CostPrice")%></td>
      <td><%=rs("StockQty")%></td>
      <td><%=rs("Remark")%></td>
      <td><%=rs("Biller")%></td>		
    </tr>
	<%
		rs.movenext
	wend
	response.Write("</table>")
  rs.close
  set rs=nothing
elseif showType="getInfo" then 
	set rs=server.createobject("adodb.recordset")
	sql="select * from Dining_Material where MaterialName='"&request("InfoID")&"'"
	rs.open sql,connzxpt,1,1
	if rs.eof then
		response.Write("0")
	else
		response.Write("1")
	end if
  rs.close
  set rs=nothing
end if
 %>
