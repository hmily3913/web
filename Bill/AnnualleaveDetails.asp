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
      datafrom=" Bill_Annualleave "
  dim datawhere'数据条件
  dim i'用于循环的整数
	if (Instr(session("AdminPurviewFLW"),"|216.3,")>0 and Depart="KD01.0001.0001") or Instr(session("AdminPurviewFLW"),"|216.2,")>0 then
    datawhere=" where left(Department,9)='"&left(Depart,9)&"' "
	elseif Instr(session("AdminPurviewFLW"),"|216.3,")>0 and Depart="KD01.0001.0010" then
		datawhere=" where (Department='"&Depart&"' or Department='KD01.0001.0018' or BillerID='"&UserName&"') "
	elseif Instr(session("AdminPurviewFLW"),"|216.3,")>0 and Depart="KD01.0005.0001" then
		datawhere=" where (left(Department,9)='"&left(Depart,9)&"' or Department='KD01.0001.0009' or BillerID='"&UserName&"') "
	else
		datawhere=" where (Department='"&Depart&"' or BillerID='"&UserName&"') "
	end if
	Dim searchterm,searchcols
	
	if Request.Form("query") <> "" then
	searchterm = Request.Form("query")
	searchcols = Request.Form("qtype")
	if isnumeric(searchterm) then
	datawhere = datawhere&" and " & searchcols & " = " & searchterm & " "
	else
	datawhere = datawhere&" and " & searchcols & " LIKE '%" & searchterm & "%' "
	end if
	else
	datawhere = datawhere&" and CancelFlag=0 "
	End if
	datawhere=datawhere&Session("AllMessage45")&Session("AllMessage46")&Session("AllMessage47")
	session.contents.remove "AllMessage45"
	session.contents.remove "AllMessage46"
	session.contents.remove "AllMessage47"
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
				tempstr="主管审核"
			elseif rs("CheckFlag")=2 then
				tempstr="已实施"
			end if
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("RegisterID")%>","<%=rs("Register")%>","<%=rs("Department")%>","<%=rs("Departmentname")%>","<%=rs("Position")%>","<%=rs("StartDate")%>","<%=rs("StartTime")%>","<%=rs("EndDate")%>","<%=rs("EndTime")%>","<%=rs("TotalHour")%>","<%=tempstr%>","<%=rs("CancelFlag")%>","<%=JsonStr(rs("Reason"))%>","<%=rs("Biller")%>","<%=rs("BillDate")%>","<%=rs("Checker")%>","<%=rs("HRer")%>"]}
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
		sql="select * from Bill_Annualleave"
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
		rs("SRC_SNum")=Request("SRC_SNum")
		rs("SRC_SNumD")=Request("SRC_SNumD")
		rs("StartDate")=Request("StartDate")
		rs("StartTime")=Request("StartTime")
		rs("EndDate")=Request("EndDate")
		rs("EndTime")=Request("EndTime")
		rs("TotalHour")=Request("TotalHour")
		rs.update
		rs.close
		set rs=nothing 
		response.write "###"
  elseif detailType="Edit"  then
		set rs = server.createobject("adodb.recordset")
		sql="select * from Bill_Annualleave where SerialNum="&request("SerialNum")
		rs.open sql,connzxpt,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if rs("CheckFlag")>0 then
			if Instr(session("AdminPurviewFLW"),"|216.3,")>0 then
				rs("BillerID")=UserName
				rs("Biller")=AdminName
				rs("BillDate")=now()
				rs("Register")=Request("Register")
				rs("RegisterID")=Request("RegisterID")
				rs("Position")=Request("Position")
				rs("Department")=Request("Department")
				rs("Departmentname")=Request("Departmentname")
				rs("Reason")=Request("Reason")
				rs("SRC_SNum")=Request("SRC_SNum")
				rs("SRC_SNumD")=Request("SRC_SNumD")
				rs("StartDate")=Request("StartDate")
				rs("StartTime")=Request("StartTime")
				rs("EndDate")=Request("EndDate")
				rs("EndTime")=Request("EndTime")
				rs("TotalHour")=Request("TotalHour")
				rs.update
				rs.close
				set rs=nothing 
				response.write "###"
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
		rs("SRC_SNum")=Request("SRC_SNum")
		rs("SRC_SNumD")=Request("SRC_SNumD")
		rs("StartDate")=Request("StartDate")
		rs("StartTime")=Request("StartTime")
		rs("EndDate")=Request("EndDate")
		rs("EndTime")=Request("EndTime")
		rs("TotalHour")=Request("TotalHour")
		rs.update
		rs.close
		set rs=nothing 
		response.write "###"
	elseif detailType="Check" then
		set rs = server.createobject("adodb.recordset")
			SerialNum=request("SerialNum")
		sql="select * from Bill_Annualleave where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
			if rs("CancelFlag") then
				response.Write("此单据已作废，不允许审核！")
				response.End()
			end if
		if rs("CheckFlag")=0 then
			dim rs2
			set rs2=connk3.Execute("select  a.FNumber,a.Name from HM_Employees a,HM_EmployeesAddInfo b where b.EM_ID = a.EM_ID and a.Status=1 and hrms_userField_18=1 and a.FNumber='"&rs("RegisterID")&"'")
			if rs2.eof and rs2.bof then
				if Instr(session("AdminPurviewFLW"),"|216.1,")=0 then
					response.write ("你没有权限进行当前操作！")
					response.end
				end if
				if rs("Department")<>Depart then
					response.write ("只能审核本部门的单据！")
					response.end
				end if
				rs("CheckerID")=UserName
				rs("Checker")=AdminName
				rs("CheckDate")=now()
				rs("CheckFlag")=1
				rs("CancelFlag")=request("operattext")
			else
				if Instr(session("AdminPurviewFLW"),"|216.2,")=0 then
					response.write ("你没有权限进行当前操作！")
					response.end
				end if
				rs("CheckerID")=UserName
				rs("Checker")=AdminName
				rs("CheckDate")=now()
				rs("CheckFlag")=1
				rs("CancelFlag")=request("operattext")
			end if
			rs2.close
			set rs2=nothing 
		elseif rs("CheckFlag")=1 then
			if Instr(session("AdminPurviewFLW"),"|216.3,")>0 then
				rs("HrerID")=UserName
				rs("Hrer")=AdminName
				rs("HrDate")=now()
				rs("CheckFlag")=2
				rs("CancelFlag")=request("operattext")
				if request("operattext")="0" then
					connzxpt.Execute("update Bill_OvertimeDetails set SurpluHour=SurpluHour+"&rs("TotalHour")&" where SerialNumD="&rs("SRC_SNumD"))
					for nnn=0 to datediff("d",rs("StartDate"),rs("EndDate"))
						dim stime,etime
						stime=" 00:00:00.000"
						if rs("StartTime")<>"" and nnn=0 then stime=" "&rs("StartTime")&":00:000"
						etime=" 23:59:00.000"
						if rs("EndTime")<>"" and nnn=datediff("d",rs("StartDate"),rs("EndDate")) then etime=" "&rs("EndTime")&":00:000"
						connkq.Execute("insert into USER_SPEDAY select userid,'"&dateadd("d",nnn,rs("StartDate"))&stime&"','"&dateadd("d",nnn,rs("StartDate"))&etime&"',4,'资讯平台调休单审核通过，单号："&rs("SerialNum")&"',getdate(),'Bill_Annualleave',"&rs("SerialNum")&" from USERINFO where ssn='"&rs("RegisterID")&"'")
					next
				end if
			else
				response.write ("你没有权限进行当前操作！")
				response.end
			end if
		end if
			rs.movenext
		wend
		rs.update
		rs.close
		set rs=nothing 
		response.write("操作成功")
  elseif detailType="Delete" then
    SerialNum=request("SerialNum")
		sql="select * from Bill_Annualleave where CheckFlag>0 and SerialNum in ("&SerialNum&")"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		if rs.eof then
			connzxpt.Execute("Delete from Bill_Annualleave where SerialNum in ("&SerialNum&")")
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
  if detailType="RegisterID" then
    InfoID=request("InfoID")
		sql="select a.*,b.部门名称 from [N-基本资料单头] a,[G-部门资料表] b where a.部门别=b.部门代号 and a.员工代号 like '%"&InfoID&"%' "
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
		sql="select a.*,b.StartDate as OStartDate,b.StartTime as OStartTime,b.EndDate as OEndDate,b.EndTime as OEndTime,b.ActualHour as OTotalHour,b.SurpluHour from Bill_Annualleave a,Bill_OvertimeDetails b where datediff(d,dateadd(m,4,b.StartDate),getdate())<=0 and a.SRC_SNumD=b.SerialNumD and a.SerialNum="&InfoID
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
  elseif detailType="OverInfo" then
		sql="select b.* from Bill_Overtime a,Bill_OvertimeDetails b where b.Overer='"&request("InfoID")&"' and b.SNum=a.Serialnum and a.CancelFlag=0 and a.CheckFlag=4 and b.ActualHour>b.SurpluHour"
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
	end if
elseif showType="Export" then 
		sql="select * from Bill_Annualleave where Startdate>='"&request("SDate")&"' and Startdate<='"&request("EDate")&"' order by Startdate"
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
