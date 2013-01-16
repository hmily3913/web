<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
if Instr(session("AdminPurviewFLW"),"|213,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
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
  dim page'页码
      page=clng(request("page"))
  dim idCount'记录总数
  dim pages'每页条数
      pages=request("rp")
  dim pagec'总页数

  dim datafrom'数据表名
      datafrom=" Bill_LeaveApplication "
  dim datawhere'数据条件
	Dim searchterm,searchcols
	if Request.Form("query") <> "" then
	searchterm = Request.Form("query")
	searchcols = request.Form("qtype")
	if searchcols = "id" then
	if isnumeric(searchterm) then
		datawhere = " WHERE " & searchcols & " = " & searchterm & ""
	else
		datawhere = " WHERE " & searchcols & " = 56465453143613645641564643156136135136561345643654"
	End if
	Else
		datawhere = " WHERE " & searchcols & " LIKE '%" & searchterm & "%'"
	End if
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
	dim ys:ys="#f7f7f7"
	if rs("CheckFlag")>0 then ys="#ffff66"
	dim CheckState:CheckState=""
	if rs("CheckFlag")="1" then
	  CheckState="主管审核"
	elseif rs("CheckFlag")="2" then
	  CheckState="人资确认"
	elseif rs("CheckFlag")="3" then
	  CheckState="副总审核"
	else
	  CheckState="待审核"
	end if
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("RegisterName")%>","<%=rs("RegDate")%>","<%=rs("Departmentname")%>","<%=rs("EntryDate")%>","<%=rs("Position")%>","<%=rs("Grade")%>","<%=rs("SubDepart")%>","<%=JsonStr(rs("LeaveReason"))%>","<%=CheckState%>"]}
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
  	if  Instr(session("AdminPurviewFLW"),"|213.1,")>0 then
			set rs = server.createobject("adodb.recordset")
			sql="select * from Bill_LeaveApplication"
			rs.open sql,connzxpt,1,3
			rs.addnew
			rs("Biller")=UserName
			rs("BillDate")=now()
			rs("Register")=Request("Register")
			rs("RegisterName")=Request("RegisterName")
			rs("RegDate")=Request("RegDate")
			rs("Department")=Request("Department")
			rs("Departmentname")=Request("Departmentname")
			rs("EntryDate")=Request("EntryDate")
			rs("Position")=Request("Position")
			rs("Grade")=Request("Grade")
			rs("SubDepart")=Request("SubDepart")
			rs("LeaveReason")=Request("LeaveReason")
			rs.update
			rs.close
			set rs=nothing 
			response.write "###"
		else
			response.write ("你没有权限进行此操作！")
			response.end
		end if
  elseif detailType="Edit"  then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_LeaveApplication where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if Instr(session("AdminPurviewFLW"),"|213.1,")=0 then
		response.write ("你没有权限进行此操作！")
		response.end
    end if
	if rs("Biller")<>UserName and rs("Register")<>UserName then
		response.write ("只能编辑自己添加的数据！")
		response.end
	end if
	if rs("CheckFlag")>0 then
		response.write ("当前状态不允许编辑，请检查！")
		response.end
	end if
			rs("Biller")=UserName
			rs("BillDate")=now()
			rs("Register")=Request("Register")
			rs("RegisterName")=Request("RegisterName")
			rs("RegDate")=Request("RegDate")
			rs("Department")=Request("Department")
			rs("Departmentname")=Request("Departmentname")
			rs("EntryDate")=Request("EntryDate")
			rs("Position")=Request("Position")
			rs("Grade")=Request("Grade")
			rs("SubDepart")=Request("SubDepart")
			rs("LeaveReason")=Request("LeaveReason")
			response.write "###"
    rs.update
	rs.close
	set rs=nothing 
  elseif detailType="Delete" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_LeaveApplication where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,1
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if Instr(session("AdminPurviewFLW"),"|213.1,")=0 then
		response.write ("你没有权限进行此操作！")
		response.end
	end if
	if rs("Biller")<>UserName and rs("Register")<>UserName then
		response.write ("只能删除本人自己添加的数据！")
		response.end
	end if
	if rs("CheckFlag")>0 then
		response.write ("已经审核不允许删除！")
		response.end
	end if
	rs.close
	set rs=nothing 
	connzxpt.Execute("Delete from Bill_LeaveApplication where SerialNum="&SerialNum)
	response.write "###"
  elseif detailType="Check" then
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select * from Bill_LeaveApplication where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if rs("CheckFlag")=0 and Instr(session("AdminPurviewFLW"),"|213.2,")>0 then
			rs("DepartReplyer")=session("AdminName")
			rs("DepartReplyDate")=now()
			rs("DepartReplyText")=Request.Form("operattext")
			rs("CheckFlag")=1
		elseif rs("CheckFlag")=1 and Instr(session("AdminPurviewFLW"),"|213.3,")>0 then
			rs("RelatedReplyer")=session("AdminName")
			rs("RelatedReplyDate")=now()
			rs("RelatedReplyText")=Request.Form("operattext")
			rs("CheckFlag")=2
		elseif rs("CheckFlag")=2 and Instr(session("AdminPurviewFLW"),"|213.4,")>0 then
			rs("VPReplyer")=session("AdminName")
			rs("VPReplyDate")=now()
			rs("VPReplyText")=Request.Form("operattext")
			rs("CheckFlag")=3
		else
			response.write ("你没有权限进行此操作或当前状态不允许此次操作！")
			response.end
		end if
		rs.update
		rs.close
		set rs=nothing 
		response.write "审核成功！"
  elseif detailType="unCheck" then
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select * from Bill_LeaveApplication where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if rs("CheckFlag")=0 then
			response.write ("此单据未审核，不允许反审核！")
			response.end
		end if
		if rs("CheckFlag")=1 and Instr(session("AdminPurviewFLW"),"|213.2,")>0 then
			rs("DepartReplyer")=session("AdminName")
			rs("DepartReplyDate")=now()
			rs("DepartReplyText")=Request.Form("operattext")
			rs("CheckFlag")=0
		elseif rs("CheckFlag")=2 and Instr(session("AdminPurviewFLW"),"|213.3,")>0 then
			rs("RelatedReplyer")=session("AdminName")
			rs("RelatedReplyDate")=now()
			rs("RelatedReplyText")=Request.Form("operattext")
			rs("CheckFlag")=1
		elseif rs("CheckFlag")=3 and Instr(session("AdminPurviewFLW"),"|213.4,")>0 then
			rs("VPReplyer")=session("AdminName")
			rs("VPReplyDate")=now()
			rs("VPReplyText")=Request.Form("operattext")
			rs("CheckFlag")=2
		else
			response.write ("你没有权限进行此操作或当前状态不允许此次操作！")
			response.end
		end if
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
	sql="select a.*,b.部门名称 from [N-基本资料单头] a,[G-部门资料表] b where a.部门别=b.部门代号 and a.员工代号='"&InfoID&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
    if rs.bof and rs.eof then
        response.write ("员工编号不存在！")
        response.end
	else
		response.write(rs("员工代号")&"###"&rs("姓名")&"###"&rs("部门别")&"###"&rs("部门名称")&"###"&rs("性别")&"###"&rs("工作岗位")&"###"&rs("到职日")&"###"&rs("职等"))
	end if
	rs.close
	set rs=nothing 
  elseif detailType="SerialNum" then
    InfoID=request("InfoID")
	sql="select * from Bill_LeaveApplication where SerialNum="&InfoID
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
    if rs.bof and rs.eof then
        response.write ("对应单据不存在，请检查！")
        response.end
	else
	  response.write "{""Info"":""###"",""fieldValue"":[{"
	  for i=0 to rs.fields.count-2
	    if IsNull(rs.fields(i).value) then
		response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
		else
		response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
		end if
	  next
	    if IsNull(rs.fields(i).value) then
		response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&"""}]}")
		else
		response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&"""}]}")
		end if
	end if
	rs.close
	set rs=nothing 
  end if
end if
 %>
