<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurview"),"|308,")=0 then 
'  response.write "{""page"":""1"",""total"":""1"",""rows"":[{""id"":""0"",""cell"":["
'  response.write ("""<font color='red')>你不具有该管理模块的操作权限，请返回！""]}]}")
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
      datafrom=" manusys_InnerProductPrice "
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
		dim shenhe
    do until rs.eof'填充数据到表格'
		shenhe="未审核"
		if rs("CheckFlag")=1 then shenhe="已审核"
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("jiagong")%>","<%=JsonStr(rs("Chengpin"))%>","<%=rs("Gongxu")%>","<%=rs("Unit")%>","<%=rs("K3Price")%>","<%=rs("Cailiao")%>","<%=rs("Dianfei")%>","<%=rs("Heji")%>","<%=rs("Bili")%>","<%=rs("Lirun")%>","<%=rs("Zongji")%>","<%=JsonStr(rs("Remark"))%>","<%=shenhe%>","<%=rs("Biller")%>","<%=rs("Checker")%>"]}
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
  	if  Instr(session("AdminPurview"),"|308.1,")>0 then
		for   i=2   to   Request.form("SerialNum").count
			if Request.form("DeleteFlag")(i)="0" and Request.form("Gongxu")(i)<>"" then
				sql="select * from manusys_InnerProductPrice where jiagong='"&Request.form("jiagong")(i)&"' and Chengpin='"&Request.form("Chengpin")(i)&"' and Gongxu='"&Request.form("Gongxu")(i)&"'"
				set rs=server.createobject("adodb.recordset")
				rs.open sql,connzxpt,1,1
				if rs.eof then
					connzxpt.Execute("insert into manusys_InnerProductPrice (jiagong,Chengpin,Gongxu,Unit,K3Price,Cailiao,Dianfei,Heji,Bili,Lirun,Zongji,Remark,BillerID,Biller,BillDate) values ('"&Request.form("jiagong")(i)&"','"&Request.form("Chengpin")(i)&"','"&Request.form("Gongxu")(i)&"','"&Request.form("Unit")(i)&"','"&Request.form("K3Price")(i)&"','"&Request.form("Cailiao")(i)&"','"&Request.form("Dianfei")(i)&"','"&Request.form("Heji")(i)&"','"&Request.form("Bili")(i)&"','"&Request.form("Lirun")(i)&"','"&Request.form("Zongji")(i)&"','"&Request.form("Remark")(i)&"','"&UserName&"','"&AdminName&"','"&now()&"')")
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
		if Instr(session("AdminPurview"),"|308.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		for   i=2   to   Request.form("SerialNum").count
			if Request.form("DeleteFlag")(i)="1" and Request.Form("SerialNum")(i)<>"" then
				connzxpt.Execute("Delete from manusys_InnerProductPrice where CheckFlag=0 and UseFlag=0 and SerialNum="&Request.Form("SerialNum")(i))
			elseif Request.Form("SerialNum")(i)<>"" then
				connzxpt.Execute("update manusys_InnerProductPrice set jiagong='"&Request.form("jiagong")(i)&"',Chengpin='"&Request.form("Chengpin")(i)&"',Gongxu='"&Request.form("Gongxu")(i)&"',Unit='"&Request.form("Unit")(i)&"',K3Price='"&Request.form("K3Price")(i)&"',Cailiao='"&Request.form("Cailiao")(i)&"',Dianfei='"&Request.form("Dianfei")(i)&"',Heji='"&Request.form("Heji")(i)&"',Bili='"&Request.form("Bili")(i)&"',Lirun='"&Request.form("Lirun")(i)&"',Zongji='"&Request.form("Zongji")(i)&"',Remark='"&Request.form("Remark")(i)&"',BillerID='"&UserName&"',Biller='"&AdminName&"',BillDate='"&now()&"' where SerialNum="&Request.Form("SerialNum")(i))
			else
				sql="select * from manusys_InnerProductPrice where jiagong='"&Request.form("jiagong")(i)&"' and Chengpin='"&Request.form("Chengpin")(i)&"' and Gongxu='"&Request.form("Gongxu")(i)&"'"
				set rs=server.createobject("adodb.recordset")
				rs.open sql,connzxpt,1,1
				if rs.eof then
					connzxpt.Execute("insert into manusys_InnerProductPrice (jiagong,Chengpin,Gongxu,Unit,K3Price,Cailiao,Dianfei,Heji,Bili,Lirun,Zongji,Remark,BillerID,Biller,BillDate) values ('"&Request.form("jiagong")(i)&"','"&Request.form("Chengpin")(i)&"','"&Request.form("Gongxu")(i)&"','"&Request.form("Unit")(i)&"','"&Request.form("K3Price")(i)&"','"&Request.form("Cailiao")(i)&"','"&Request.form("Dianfei")(i)&"','"&Request.form("Heji")(i)&"','"&Request.form("Bili")(i)&"','"&Request.form("Lirun")(i)&"','"&Request.form("Zongji")(i)&"','"&Request.form("Remark")(i)&"','"&UserName&"','"&AdminName&"','"&now()&"')")
				end if
				rs.close
				set rs=nothing
			end if
		next
		response.write "###"
  elseif detailType="Delete" then
		if Instr(session("AdminPurview"),"|308.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
    SerialNum=request("SerialNum")
		sql="select * from manusys_InnerProductPrice where (CheckFlag=1 or UseFlag=1) and SerialNum in ("&SerialNum&")"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		if rs.eof then
			connzxpt.Execute("Delete from manusys_InnerProductPrice where SerialNum in ("&SerialNum&")")
			response.write "###"
		else
			response.Write("已审核,或者已使用，不允许删除！")
			response.End()
		end if
		rs.close
		set rs=nothing
  elseif detailType="Check"  then
		if Instr(session("AdminPurview"),"|308.2,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		connzxpt.Execute("update manusys_InnerProductPrice set CheckFlag=1,CheckerID='"&UserName&"',Checker='"&AdminName&"',CheckDate='"&now()&"' where SerialNum in ("&request("SerialNum")&")")
		response.write "###"
  elseif detailType="UnCheck"  then
		if Instr(session("AdminPurview"),"|308.2,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		connzxpt.Execute("update manusys_InnerProductPrice set CheckFlag=0,CheckerID='"&UserName&"',Checker='"&AdminName&"',CheckDate='"&now()&"' where SerialNum in ("&request("SerialNum")&")")
		response.write "###"
  end if
elseif showType="Export" then 
	set rs=server.createobject("adodb.recordset")
	sql="select * from manusys_InnerProductPrice"
	rs.open sql,connzxpt,1,1
	%>
  <table border="1">
  	<tr>
			<td width="6%" bgcolor="#8DB5E9"><strong>加工方</strong></td>
			<td width="8%" bgcolor="#8DB5E9"><strong>成品品号</strong></td>
			<td width="10%" bgcolor="#8DB5E9"><strong>工序名称</strong></td>
			<td width="4%" bgcolor="#8DB5E9"><strong>单位</strong></td>
			<td width="6%" bgcolor="#8DB5E9"><strong>K3单价</strong></td>
			<td width="6%" bgcolor="#8DB5E9"><strong>材料费用</strong></td>
			<td width="6%" bgcolor="#8DB5E9"><strong>电费</strong></td>
			<td width="6%" bgcolor="#8DB5E9"><strong>合计单价</strong></td>
			<td width="6%" bgcolor="#8DB5E9"><strong>相关比例</strong></td>
			<td width="6%" bgcolor="#8DB5E9"><strong>利润率</strong></td>
			<td width="6%" bgcolor="#8DB5E9"><strong>总计单价</strong></td>
			<td width="10%" bgcolor="#8DB5E9"><strong>备注</strong></td>
    </tr>
	<%
	while (not rs.eof)
	%>
	  <tr>
      <td><%=rs("jiagong")%></td>
      <td><%=rs("Chengpin")%></td>
      <td><%=rs("Gongxu")%></td>
      <td><%=rs("Unit")%></td>
      <td><%=rs("K3Price")%></td>
      <td><%=rs("Cailiao")%></td>
      <td><%=rs("Dianfei")%></td>
      <td><%=rs("Heji")%></td>
      <td><%=rs("Bili")%></td>
      <td><%=rs("Lirun")%></td>
      <td><%=rs("Zongji")%></td>
      <td><%=rs("Remark")%></td>
    </tr>
	<%
		rs.movenext
	wend
	response.Write("</table>")
  rs.close
  set rs=nothing
end if
 %>
