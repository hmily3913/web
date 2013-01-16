<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurview"),"|310,")=0 then 
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
      datafrom=" manusys_OutProduct "
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
		 if request("sd")<>"" then datawhere=datawhere&" and datediff(d,BillDate,'"&request("sd")&"')<=0 "
		 if request("ed")<>"" then datawhere=datawhere&" and datediff(d,BillDate,'"&request("ed")&"')>=0 "
		 if request("wt")<>"" then datawhere=datawhere&" and Weituo like '%"&request("wt")&"%' "
		 if request("jg")<>"" then datawhere=datawhere&" and jiagong='"&request("jg")&"' "
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
		dim checktag
    do until rs.eof'填充数据到表格'
		checktag="未审"
		if rs("CheckFlag")="1" then checktag="已审"
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("Weituo")%>","<%=rs("jiagong")%>","<%=rs("Chengpin")%>","<%=rs("Gongxu")%>","<%=rs("Unit")%>","<%=rs("Danjia")%>","<%=rs("Shuliang")%>","<%=rs("Yingshou")%>","<%=rs("Shishou")%>","<%=checktag%>","<%=rs("Remark")%>","<%=rs("Biller")%>","<%=rs("BillDate")%>","<%=rs("Checker")%>"]}
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
  	if  Instr(session("AdminPurview"),"|310.1,")>0 then
			for   i=2   to   Request.form("SerialNum").count
				if (Request.form("DeleteFlag")(i)="0" or Request.form("DeleteFlag")(i)="") and Request.form("GongxuID")(i)<>"" then
					connzxpt.Execute("insert into manusys_OutProduct (Shishou,jiagong,Weituo,Shuliang,Danjia,Yingshou,Remark,BillerID,Biller,BillDate,Chengpin,Gongxu,Unit) values ("&Request.form("Shishou")(i)&",'"&Request.form("jiagong")(i)&"','"&Request.form("Weituo")(i)&"','"&Request.form("Shuliang")(i)&"','"&Request.form("Danjia")(i)&"','"&Request.form("Yingshou")(i)&"','"&Request.form("Remark")(i)&"','"&UserName&"','"&AdminName&"','"&now()&"','"&Request.form("Chengpin")(i)&"','"&Request.form("Gongxu")(i)&"','"&Request.form("Unit")(i)&"')")
				end if
			next
			response.write "###"
		else
			response.write ("你没有权限进行此操作！")
			response.end
		end if
  elseif detailType="Edit"  then
		if Instr(session("AdminPurview"),"|310.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		for   i=2   to   Request.form("SerialNum").count
			if Request.form("DeleteFlag")(i)="1" and Request.Form("SerialNum")(i)<>"" then
				connzxpt.Execute("Delete from manusys_OutProduct where SerialNum="&Request.Form("SerialNum")(i))
			elseif Request.Form("SerialNum")(i)<>"" then
				connzxpt.Execute("update manusys_OutProduct set Shishou='"&Request.form("Shishou")(i)&"',Weituo='"&Request.form("Weituo")(i)&"',jiagong='"&Request.form("jiagong")(i)&"',Shuliang='"&Request.form("Shuliang")(i)&"',Danjia='"&Request.form("Danjia")(i)&"',Yingshou='"&Request.form("Yingshou")(i)&"',Remark='"&Request.form("Remark")(i)&"',BillerID='"&UserName&"',Biller='"&AdminName&"',BillDate='"&now()&"',Chengpin='"&Request.form("Chengpin")(i)&"',Gongxu='"&Request.form("Gongxu")(i)&"',Unit='"&Request.form("Unit")(i)&"' where SerialNum="&Request.Form("SerialNum")(i))
			else
					connzxpt.Execute("insert into manusys_OutProduct (Shishou,jiagong,Weituo,Shuliang,Danjia,Yingshou,Remark,BillerID,Biller,BillDate,Chengpin,Gongxu,Unit) values ("&Request.form("Shishou")(i)&",'"&Request.form("jiagong")(i)&"','"&Request.form("Weituo")(i)&"','"&Request.form("Shuliang")(i)&"','"&Request.form("Danjia")(i)&"','"&Request.form("Yingshou")(i)&"','"&Request.form("Remark")(i)&"','"&UserName&"','"&AdminName&"','"&now()&"','"&Request.form("Chengpin")(i)&"','"&Request.form("Gongxu")(i)&"','"&Request.form("Unit")(i)&"')")
			end if
		next
		response.write "###"
  elseif detailType="Delete" then
		if Instr(session("AdminPurview"),"|310.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
    SerialNum=request("SerialNum")
		sql="select * from manusys_OutProduct where CheckFlag=1 and SerialNum in ("&SerialNum&")"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		if rs.eof then
			connzxpt.Execute("Delete from manusys_OutProduct where SerialNum in ("&SerialNum&")")
			response.write "###"
		else
			response.Write("已审核不允许删除！")
			response.End()
		end if
  elseif detailType="Check"  then
		SerialNum=request("SerialNum")
		if Instr(session("AdminPurview"),"|310.2,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		sql="select * from manusys_OutProduct where SerialNum in ("&SerialNum&")"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
			if rs("CheckFlag")=0 then
				rs("Checker")=AdminName
				rs("CheckerID")=UserName
				rs("CheckDate")=now()
				rs("CheckFlag")=1
			end if		
			rs.movenext
		wend
		rs.update
		rs.close
		set rs=nothing 
		response.write "###"
  elseif detailType="unCheck"  then
		if Instr(session("AdminPurview"),"|310.2,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		SerialNum=request("SerialNum")
		connzxpt.Execute("update manusys_OutProduct set CheckFlag=0 where SerialNum in ("&SerialNum&")")
		response.write "###"
  end if
elseif showType="Export" then 
%>
 <table width="100%" border="1" cellpadding="3" cellspacing="1">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><strong>委托方</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>加工方</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>成品名称</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>工序名称</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>单位</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>单价</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>数量</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>应收金额</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>实收金额</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>审核</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>备注</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>登记人</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>登记日期</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>审核人</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>审核日期</strong></td>
  </tr>
 <%
    sql="select * from manusys_OutProduct where 1=1 "
		 if request("sd")<>"" then sql=sql&" and datediff(d,BillDate,'"&request("sd")&"')<=0 "
		 if request("ed")<>"" then sql=sql&" and datediff(d,BillDate,'"&request("ed")&"')>=0 "
		 if request("wt")<>"" then sql=sql&" and Weituo like '%"&request("wt")&"%' "
		 if request("jg")<>"" then sql=sql&" and jiagong='"&request("jg")&"' "
		 sql=sql&" order by SerialNum desc "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,0,1
    while(not rs.eof)'填充数据到表格
		checktag="未审"
		if rs("CheckFlag")="1" then checktag="已审"
	  Response.Write "<tr>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Weituo")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("jiagong")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Chengpin")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Gongxu")&"</td>"
      Response.Write "<td nowrap>"&rs("Unit")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Danjia")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Shuliang")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Yingshou")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Shishou")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&checktag&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Remark")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Biller")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("BillDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Checker")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("CheckDate")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
	  rs2.close
	  set rs2=nothing
  %>
  </table>
<% 
end if
 %>
