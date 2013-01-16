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
      datafrom=" Bill_Leave "
  dim datawhere'数据条件
  dim i'用于循环的整数
	if (Instr(session("AdminPurviewFLW"),"|217.3,")>0 and Depart="KD01.0001.0001") or Instr(session("AdminPurviewFLW"),"|217.2,")>0 then
    datawhere=" where left(Department,9)='"&left(Depart,9)&"' "
	elseif Instr(session("AdminPurviewFLW"),"|217.3,")>0 and Depart="KD01.0005.0001" then
		datawhere=" where (left(Department,9)='"&left(Depart,9)&"' or Department='KD01.0001.0009' or BillerID='"&UserName&"') "
	else
		datawhere=" where (Department='"&Depart&"' or BillerID='"&UserName&"') "
	end if
	Dim searchterm,searchcols
	if request("AllQuery")<>"" then
		datawhere=datawhere&" and "&request("AllQuery")
	else
		datawhere=datawhere&" and CancelFlag=0 "
	end if

	datawhere=datawhere&Session("AllMessage48")&Session("AllMessage49")&Session("AllMessage50")&Session("AllMessage56")
	session.contents.remove "AllMessage48"
	session.contents.remove "AllMessage49"
	session.contents.remove "AllMessage50"
	session.contents.remove "AllMessage56"
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
			dim tempstr
			if rs("CheckFlag")=0 then
				tempstr="未审核"
			elseif rs("CheckFlag")=1 then
				tempstr="工段长审核"
			elseif rs("CheckFlag")=2 then
				tempstr="主管审核"
			elseif rs("CheckFlag")=3 then
				tempstr="已审批"
			elseif rs("CheckFlag")=4 then
				tempstr="已实施"
			end if
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("RegisterID")%>","<%=rs("Register")%>","<%=rs("Department")%>","<%=rs("Departmentname")%>","<%=rs("Position")%>","<%=rs("StartDate")%>","<%=rs("StartTime")%>","<%=rs("EndDate")%>","<%=rs("EndTime")%>","<%=rs("TotalDay")%>","<%=rs("TotalHour")%>","<%=tempstr%>","<%=rs("CancelFlag")%>","<%=JsonStr(rs("Reason"))%>","<%=rs("Biller")%>","<%=rs("BillDate")%>","<%=rs("Checker2")%>","<%=rs("Checker")%>","<%=rs("Approvaler")%>","<%=rs("Hrer")%>"]}
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
		set rs = server.createobject("adodb.recordset")
		sql="select * from Bill_Leave"
		rs.open sql,connzxpt,1,3
		rs.addnew
		rs("BillerID")=UserName
		rs("Biller")=AdminName
		rs("BillDate")=now()
		rs("Register")=Request("Register")
		rs("RegisterID")=Request("RegisterID")
		rs("Position")=Request("Position")
		rs("Department")=Request("Department")
		rs("Departmentname")=Request("Departmentname")
		rs("Reason")=Request("Reason")
		rs("AgentID")=Request("AgentID")
		rs("Agenter")=Request("Agenter")
		rs("StartDate")=Request("StartDate")
		rs("StartTime")=Request("StartTime")
		rs("EndDate")=Request("EndDate")
		rs("EndTime")=Request("EndTime")
		rs("LeaveType")=Request("LeaveType")
		rs("TotalDay")=Request("TotalDay")
		rs("TotalHour")=Request("TotalHour")
		rs("Grade")=Request("Grade")
		rs("SalaryType")=Request("SalaryType")
		rs.update
		rs.close
		set rs=nothing 
		response.write "###"
  elseif detailType="Edit"  then
		set rs = server.createobject("adodb.recordset")
		sql="select * from Bill_Leave where SerialNum="&request("SerialNum")
		rs.open sql,connzxpt,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if rs("CheckFlag")>0 then
			if Instr(session("AdminPurviewFLW"),"|217.3,")>0 then
		rs("StartDate")=Request("StartDate")
		rs("StartTime")=Request("StartTime")
		rs("EndDate")=Request("EndDate")
		rs("EndTime")=Request("EndTime")
		rs("TotalDay")=Request("TotalDay")
		rs("TotalHour")=Request("TotalHour")
		rs.update
		rs.close
		set rs=nothing 
		response.write "###"
		response.end
			else
				response.write ("当前状态不允许编辑，请检查！")
				response.end
			end if
		end if
		if rs("BillerID")<>UserName and rs("RegisterID")<>UserName then
			response.write ("只能编辑自己添加的数据！")
			response.end
		end if
		rs("BillerID")=UserName
		rs("Biller")=AdminName
		rs("BillDate")=now()
		rs("Register")=Request("Register")
		rs("RegisterID")=Request("RegisterID")
		rs("Position")=Request("Position")
		rs("Department")=Request("Department")
		rs("Departmentname")=Request("Departmentname")
		rs("Reason")=Request("Reason")
		rs("AgentID")=Request("AgentID")
		rs("Agenter")=Request("Agenter")
		rs("StartDate")=Request("StartDate")
		rs("StartTime")=Request("StartTime")
		rs("EndDate")=Request("EndDate")
		rs("EndTime")=Request("EndTime")
		rs("LeaveType")=Request("LeaveType")
		rs("TotalDay")=Request("TotalDay")
		rs("TotalHour")=Request("TotalHour")
		rs("Grade")=Request("Grade")
		rs("SalaryType")=Request("SalaryType")
		rs.update
		rs.close
		set rs=nothing 
		response.write "###"
	elseif detailType="Check" then
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select * from Bill_Leave where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
			if rs("CancelFlag") then
				response.Write("此单据已作废，不允许审核！")
				response.End()
			end if
			if rs("CheckFlag")>2 then
				response.Write("此单据不允许进行此操作！")
				response.End()
			end if
			if rs("CheckFlag")=0 and (rs("Department")="KD01.0001.0018" or rs("Department")="KD01.0001.0010" or rs("Department")="KD01.0001.0007" or rs("Department")="KD01.0001.0008" or rs("Department")="KD01.0001.0019" or rs("Department")="KD01.0004.0001") and Instr(session("AdminPurviewFLW"),"|217.4,")>0 and cdbl(rs("Grade"))<5 then
				rs("CheckerID2")=UserName
				rs("Checker2")=AdminName
				rs("CheckDate2")=now()
				rs("CheckFlag")=1
				rs("CancelFlag")=request("operattext")
			elseif Instr(session("AdminPurviewFLW"),"|217.1,")>0 then
				rs("CheckerID")=UserName
				rs("Checker")=AdminName
				rs("CheckDate")=now()
				rs("CheckFlag")=2
				rs("CancelFlag")=request("operattext")
			else 
				response.Write("你没有权限进行此操作！")
				response.End()
			end if
			rs.movenext
		wend
		rs.update
		rs.close
		set rs=nothing 
		response.write("审核成功！")
	elseif detailType="Approval" then
		set rs = server.createobject("adodb.recordset")
			SerialNum=request("SerialNum")
		sql="select * from Bill_Leave where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
			if rs("CancelFlag") then
				response.Write("此单据已作废，不允许审核！")
				response.End()
			end if
			if rs("CheckFlag")>3 then
				response.Write("此单据不允许进行此操作！")
				response.End()
			end if
			if Instr(session("AdminPurviewFLW"),"|217.2,")=0 then
				response.write ("你没有权限进行当前操作！")
				response.end
			end if
			rs("ApprovalerID")=UserName
			rs("Approvaler")=AdminName
			rs("ApprovalDate")=now()
			rs("CheckFlag")=3
			rs("CancelFlag")=request("operattext")
			rs.movenext
		wend
		rs.update
		rs.close
		set rs=nothing 
		response.write("审批成功！")
	elseif detailType="Hr" then
		set rs = server.createobject("adodb.recordset")
			SerialNum=request("SerialNum")
		sql="select * from Bill_Leave where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
			if rs("CancelFlag") then
				response.Write("此单据已作废，不允许审核！")
				response.End()
			end if
			if Instr(session("AdminPurviewFLW"),"|217.3,")=0 then
				response.write ("你没有权限进行当前操作！")
				response.end
			end if
			if rs("CheckFlag")=1 and cdbl(rs("TotalDay"))<7 and (rs("SalaryType")="3.计件" or rs("SalaryType")="2.分厂月薪") then
				rs("HrerID")=UserName
				rs("Hrer")=AdminName
				rs("HrDate")=now()
				rs("CheckFlag")=4
				rs("CancelFlag")=request("operattext")
				rs.update
				if request("operattext")="0" then
					if rs("LeaveType")="病假" then
						qjlb=1
					elseif rs("LeaveType")="探亲假" then
						qjlb=3
					else
						qjlb=2
					end if
					for nnn=0 to datediff("d",rs("StartDate"),rs("EndDate"))
						stime=" 00:00:00.000"
						if rs("StartTime")<>"" and nnn=0 then stime=" "&rs("StartTime")&":00:000"
						etime=" 23:59:00.000"
						if rs("EndTime")<>"" and nnn=datediff("d",rs("StartDate"),rs("EndDate")) then etime=" "&rs("EndTime")&":00:000"
						connkq.Execute("insert into USER_SPEDAY select userid,'"&dateadd("d",nnn,rs("StartDate"))&stime&"','"&dateadd("d",nnn,rs("StartDate"))&etime&"',"&qjlb&",'资讯平台请假单审核通过，单号："&rs("SerialNum")&"',getdate(),'Bill_Leave',"&rs("SerialNum")&" from USERINFO where ssn='"&rs("RegisterID")&"'")
					next
				end if
			elseif rs("CheckFlag")=2 then
				dim rs2,stime,etime
				set rs2=connk3.Execute("select  a.FNumber,a.Name from HM_Employees a,HM_EmployeesAddInfo b where b.EM_ID = a.EM_ID and a.Status=1 and hrms_userField_18=1 and a.FNumber='"&rs("RegisterID")&"'")
				if rs2.eof and rs2.bof and ((cdbl(rs("TotalDay"))<7 and rs("SalaryType")="1.行政月薪") or (cdbl(rs("TotalDay"))<15 and (rs("SalaryType")="3.计件" or rs("SalaryType")="2.分厂月薪"))) or left(rs("Department"),9)="KD01.0005" then
					rs("HrerID")=UserName
					rs("Hrer")=AdminName
					rs("HrDate")=now()
					rs("CheckFlag")=4
					rs("CancelFlag")=request("operattext")
					rs.update
					if request("operattext")="0" then
						if rs("LeaveType")="病假" then
							qjlb=1
						elseif rs("LeaveType")="探亲假" then
							qjlb=3
						else
							qjlb=2
						end if
						for nnn=0 to datediff("d",rs("StartDate"),rs("EndDate"))
							stime=" 00:00:00.000"
							if rs("StartTime")<>"" and nnn=0 then stime=" "&rs("StartTime")&":00:000"
							etime=" 23:59:00.000"
							if rs("EndTime")<>"" and nnn=datediff("d",rs("StartDate"),rs("EndDate")) then etime=" "&rs("EndTime")&":00:000"
							connkq.Execute("insert into USER_SPEDAY select userid,'"&dateadd("d",nnn,rs("StartDate"))&stime&"','"&dateadd("d",nnn,rs("StartDate"))&etime&"',"&qjlb&",'资讯平台请假单审核通过，单号："&rs("SerialNum")&"',getdate(),'Bill_Leave',"&rs("SerialNum")&" from USERINFO where ssn='"&rs("RegisterID")&"'")
						next
					end if
				else
					response.Write("此单据需审批才能执行！")
					response.End()
				end if
			elseif rs("CheckFlag")=3 then
				rs("HrerID")=UserName
				rs("Hrer")=AdminName
				rs("HrDate")=now()
				rs("CheckFlag")=4
				rs("CancelFlag")=request("operattext")
				rs.update
				if request("operattext")="0" then
					if rs("LeaveType")="病假" then
						qjlb=1
					elseif rs("LeaveType")="探亲假" then
						qjlb=3
					else
						qjlb=2
					end if
					for nnn=0 to datediff("d",rs("StartDate"),rs("EndDate"))
						stime=" 00:00:00.000"
						if rs("StartTime")<>"" and nnn=0 then stime=" "&rs("StartTime")&":00:000"
						etime=" 23:59:00.000"
						if rs("EndTime")<>"" and nnn=datediff("d",rs("StartDate"),rs("EndDate")) then etime=" "&rs("EndTime")&":00:000"
						connkq.Execute("insert into USER_SPEDAY select userid,'"&dateadd("d",nnn,rs("StartDate"))&stime&"','"&dateadd("d",nnn,rs("StartDate"))&etime&"',"&qjlb&",'资讯平台请假单审核通过，单号："&rs("SerialNum")&"',getdate(),'Bill_Leave',"&rs("SerialNum")&" from USERINFO where ssn='"&rs("RegisterID")&"'")
					next
				end if
			else
				response.Write("此单据已执行，不允许进行此操作！")
				response.End()
			end if
			rs.movenext
		wend
		rs.close
		set rs=nothing 
		response.write("实施成功！")
  elseif detailType="Delete" then
    SerialNum=request("SerialNum")
		sql="select * from Bill_Leave where CheckFlag>0 and SerialNum in ("&SerialNum&")"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		if rs.eof then
			connzxpt.Execute("Delete from Bill_Leave where SerialNum in ("&SerialNum&")")
			response.write "###"
		else
			response.Write("已审核不允许删除！")
			response.End()
		end if
		rs.close
		set rs=nothing
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="RegisterID" or detailType="AgentID" then
    InfoID=request("InfoID")
		sql="select a.*,b.部门名称 from [N-基本资料单头] a,[G-部门资料表] b where a.部门别=b.部门代号 and a.员工代号 like '%"&InfoID&"%' "
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
				response.write ("员工编号不存在！")
				response.end
		else
			response.write(rs("员工代号")&"###"&rs("姓名")&"###"&rs("部门别")&"###"&rs("部门名称")&"###"&rs("性别")&"###"&rs("工作岗位")&"###"&rs("到职日")&"###"&rs("职等")&"###"&rs("薪资别"))
		end if
		rs.close
		set rs=nothing 
  elseif detailType="SerialNum" then
    InfoID=request("InfoID")
		sql="select * from Bill_Leave where SerialNum="&InfoID
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
elseif showType="Export" then 
		sql="select * from Bill_Leave where Startdate>='"&request("SDate")&"' and Startdate<='"&request("EDate")&"' order by Startdate"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		%>
    <table border="1">
    <tr>
		<%
		for i=0 to rs.fields.count-1
			response.Write("<td>"&rs.fields(i).name & "</td>")
		next
		response.Write("</tr>")
		response.Write("<tr>")
		while(not rs.eof)
			for i=0 to rs.fields.count-1
				response.write ("<td>"&rs.fields(i).value&"</td>")
			next
			response.write ("</tr>")
			rs.movenext
		wend
		response.Write("</table>")
		rs.close
		set rs=nothing 
end if
 %>
