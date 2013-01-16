<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurviewFLW"),"|214,")=0 then 
'  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
'  response.end
'end if
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
      datafrom=" Attendance_Travel "
  dim datawhere'数据条件
	if Instr(session("AdminPurviewFLW"),"|701.3,")>0 or Instr(session("AdminPurviewFLW"),"|701.2,")>0 then
    datawhere=" where Left(Department,9)='"&left(Depart,9)&"' "
	else
		datawhere=" where (Department='"&Depart&"' or Biller='"&AdminName&"') "
	end if
	Dim searchterm,searchcols
	if Request.Form("query") <> "" then
	searchterm = Request.Form("query")
	searchcols = request.Form("qtype")
		if Request.Form("qtype")="RegDate" then
			datawhere = datawhere&" and datediff(d,"&searchcols&",'"&searchterm&"')=0 "
		else
			datawhere = datawhere&" and "&searchcols&" like '%"&searchterm&"%' "
		end if
	else
	datawhere = datawhere&" and CancelFlag=0 "
	End if
	datawhere = datawhere&Session("AllMessage60")&Session("AllMessage61")&Session("AllMessage62")
	session.contents.remove "AllMessage60"
	session.contents.remove "AllMessage61"
	session.contents.remove "AllMessage62"
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
	  CheckState="副总审批"
	elseif rs("CheckFlag")="3" then
	  CheckState="已执行"
	else
	  CheckState="待审核"
	end if
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("Register")%>","<%=rs("RegisterName")%>","<%=rs("RegDate")%>","<%=rs("Departmentname")%>","<%=rs("StartDate")%>","<%=rs("EndDate")%>","<%=rs("TotalDays")%>","<%=JsonStr(rs("Itinerary"))%>","<%=CheckState%>","<%=rs("CancelFlag")%>","<%=rs("Biller")%>","<%=rs("BillDate")%>","<%=rs("Checker")%>","<%=rs("VPReplyer")%>","<%=rs("Hrer")%>"]}
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
'  	if  Instr(session("AdminPurviewFLW"),"|214.1,")>0 then
			set rs = server.createobject("adodb.recordset")
			sql="select * from Attendance_Travel"
			rs.open sql,connzxpt,1,3
			rs.addnew
			rs("Biller")=AdminName
			rs("BillerID")=UserName
			rs("BillDate")=now()
			rs("Register")=Request("Register")
			rs("RegisterName")=Request("RegisterName")
			rs("RegDate")=Request("RegDate")
			rs("Department")=Request("Department")
			rs("Departmentname")=Request("Departmentname")
			rs("AgentID")=Request("AgentID")
			rs("Agenter")=Request("Agenter")
			rs("Matter")=Request("Matter")
			rs("ExpectedTarg")=Request("ExpectedTarg")
			rs("StayStandard")=Request("StayStandard")
			rs("DoublePerson")=Request("DoublePerson")
			rs("OnePerson")=Request("OnePerson")
			rs("Allowance")=Request("Allowance")
			rs("TravelAdvance")=Request("TravelAdvance")
			rs("StartDate")=Request("StartDate")
			rs("EndDate")=Request("EndDate")
			rs("StartTime")=Request("StartTime")
			rs("EndTime")=Request("EndTime")
			rs("TotalDays")=Request("TotalDays")
			rs("Itinerary")=Request("Itinerary")
			rs("Level")=Request("Level")
			rs("CityType")=Request("CityType")
			rs.update
			set rs=connzxpt.Execute("select top 1 SerialNum from Attendance_Travel order by serialnum desc")
			if rs("SerialNum")="" then
			SerialNum=10000000
			else
			SerialNum=rs("SerialNum")
			end if
			dim TravelNum:TravelNum=1
			for   i=2   to   Request.form("SerialNumD").count
				if (Request.form("DeleteFlag")(i)="0" or Request.form("DeleteFlag")(i)="") and Request.form("Accompany")(i)<>"" then
					connzxpt.Execute("insert into Attendance_TravelDetails (SNum,Accompany,AccompanyName,AcDepartment,AcDepartmentname,Position,Grade,ActEndDate,ActEndTime) values ('"&SerialNum&"','"&Request.form("Accompany")(i)&"','"&Request.form("AccompanyName")(i)&"','"&Request.form("AcDepartment")(i)&"','"&Request.form("AcDepartmentname")(i)&"','"&Request.form("Position")(i)&"','"&Request.form("Grade")(i)&"','"&Request.form("ActEndDate")(i)&"','"&Request.form("ActEndTime")(i)&"')")
					TravelNum=TravelNum+1
				end if
			next
			connzxpt.Execute("update Attendance_Travel set TravelNum="&TravelNum&" where serialnum="&SerialNum)
			rs.close
			set rs=nothing 
			response.write "保存成功！"
'		else
'			response.write ("你没有权限进行此操作！")
'			response.end
'		end if
  elseif detailType="Edit"  then
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select * from Attendance_Travel where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
'		if Instr(session("AdminPurviewFLW"),"|214.1,")=0 then
'			response.write ("你没有权限进行此操作！")
'			response.end
'			end if
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
			rs("RegDate")=Request("RegDate")
			rs("Department")=Request("Department")
			rs("Departmentname")=Request("Departmentname")
			rs("AgentID")=Request("AgentID")
			rs("Agenter")=Request("Agenter")
			rs("Matter")=Request("Matter")
			rs("ExpectedTarg")=Request("ExpectedTarg")
			rs("StayStandard")=Request("StayStandard")
			rs("DoublePerson")=Request("DoublePerson")
			rs("OnePerson")=Request("OnePerson")
			rs("Allowance")=Request("Allowance")
			rs("TravelAdvance")=Request("TravelAdvance")
			rs("StartDate")=Request("StartDate")
			rs("EndDate")=Request("EndDate")
			rs("StartTime")=Request("StartTime")
			rs("EndTime")=Request("EndTime")
			rs("TotalDays")=Request("TotalDays")
			rs("Itinerary")=Request("Itinerary")
			rs("Level")=Request("Level")
			rs("CityType")=Request("CityType")
    rs.update
		TravelNum=1
		for   i=2   to   Request.form("SerialNumD").count
			if Request.form("DeleteFlag")(i)="1" and Request.Form("SerialNumD")(i)<>"" then
				connzxpt.Execute("Delete from Attendance_TravelDetails where SerialNumD="&Request.Form("SerialNumD")(i))
			elseif Request.Form("SerialNumD")(i)<>"" then
				connzxpt.Execute("update Attendance_TravelDetails set Accompany='"&Request.form("Accompany")(i)&"',AccompanyName='"&Request.form("AccompanyName")(i)&"',AcDepartment='"&Request.form("AcDepartment")(i)&"',AcDepartmentname='"&Request.form("AcDepartmentname")(i)&"',Position='"&Request.form("Position")(i)&"',Grade='"&Request.form("Grade")(i)&"',ActEndDate='"&Request.form("ActEndDate")(i)&"',ActEndTime='"&Request.form("ActEndTime")(i)&"' where SerialNumD="&Request.Form("SerialNumD")(i))
				TravelNum=TravelNum+1
			elseif Request.form("Accompany")(i)<>"" then
					connzxpt.Execute("insert into Attendance_TravelDetails (SNum,Accompany,AccompanyName,AcDepartment,AcDepartmentname,Position,Grade,ActEndDate,ActEndTime) values ('"&SerialNum&"','"&Request.form("Accompany")(i)&"','"&Request.form("AccompanyName")(i)&"','"&Request.form("AcDepartment")(i)&"','"&Request.form("AcDepartmentname")(i)&"','"&Request.form("Position")(i)&"','"&Request.form("Grade")(i)&"','"&Request.form("ActEndDate")(i)&"','"&Request.form("ActEndTime")(i)&"')")
					TravelNum=TravelNum+1
			end if
		next
		connzxpt.Execute("update Attendance_Travel set TravelNum="&TravelNum&" where serialnum="&SerialNum)
		rs.close
		set rs=nothing 
		response.write "修改成功！"
  elseif detailType="Delete" then
'		if Instr(session("AdminPurviewFLW"),"|214.1,")=0 then
'			response.write ("你没有权限进行此操作！")
'			response.end
'		end if
		set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
		sql="select * from Attendance_Travel where CheckFlag>0 and SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,1
		if rs.bof and rs.eof then
			connzxpt.Execute("Delete from Attendance_Travel where SerialNum in ("&SerialNum&")")
			connzxpt.Execute("Delete from Attendance_TravelDetails where SNum in ("&SerialNum&")")
			response.write "删除成功！"
		else
			response.write ("已经审核不允许删除！")
			response.end
		end if
		rs.close
		set rs=nothing 
  elseif detailType="Check" then
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select * from Attendance_Travel where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
			if rs("CancelFlag") then
				response.Write("此单据已作废，不允许审核！")
				response.End()
			end if
			if rs("CheckFlag")=0 and Instr(session("AdminPurviewFLW"),"|701.1,")>0 then
				rs("Checker")=AdminName
				rs("CheckerID")=UserName
				rs("CheckDate")=now()
				rs("CheckFlag")=1
				rs("CancelFlag")=request("operattext")
			elseif rs("CheckFlag")=1 and Instr(session("AdminPurviewFLW"),"|701.2,")>0 then
				rs("VPReplyer")=AdminName
				rs("VPReplyerID")=UserName
				rs("VPReplyDate")=now()
				rs("CheckFlag")=2
				rs("CancelFlag")=request("operattext")
			elseif rs("CheckFlag")=2 and Instr(session("AdminPurviewFLW"),"|701.3,")>0 then
				rs("Hrer")=AdminName
				rs("HrerID")=UserName
				rs("HrDate")=now()
				rs("CheckFlag")=3
				for   i=2   to   Request.form("SerialNumD").count
					dim overflag
					overflag=0
					if Request.form("OverFlag")(i)="True" then overflag=1
					connzxpt.Execute("update Attendance_TravelDetails set ActEndDate='"&Request.form("ActEndDate")(i)&"',ActEndTime='"&Request.form("ActEndTime")(i)&"',OverFlag="&overflag&" where SerialNumD="&Request.Form("SerialNumD")(i))
				next
				set rs2 = server.createobject("adodb.recordset")
				sql2="select * from Attendance_TravelDetails where SNum="&rs("SerialNum")
				rs2.open sql2,connzxpt,1,1
				while(not rs2.eof)
					for nnn=0 to datediff("d",rs("StartDate"),rs2("ActEndDate"))
						stime=" 00:00:00.000"
						if rs("StartTime")<>"" and nnn=0 then stime=" "&rs("StartTime")&":00:000"
						etime=" 23:59:00.000"
						if rs2("ActEndTime")<>"" and nnn=datediff("d",rs("StartDate"),rs2("ActEndDate")) then etime=" "&rs2("ActEndTime")&":00:000"
						connkq.Execute("insert into USER_SPEDAY select userid,'"&dateadd("d",nnn,rs("StartDate"))&stime&"','"&dateadd("d",nnn,rs("StartDate"))&etime&"',7,'资讯平台出差单审核通过，单号："&rs("SerialNum")&"',getdate(),'Attendance_Travel',"&rs("SerialNum")&" from USERINFO where ssn='"&rs2("Accompany")&"'")
						
						if Weekday(dateadd("d",nnn,rs("StartDate")))=1 and rs2("OverFlag") then
							dim OSerialNum:OSerialNum=getBillNo("Bill_Overtime",3,date())
							connzxpt.Execute("insert into Bill_Overtime select '"&rs("StartDate")&"','"&rs("Register")&"','"&rs("RegisterName")&"','"&rs("Department")&"','"&rs("Departmentname")&"','出差加班','"&rs("SerialNum")&"','"&rs("Departmentname")&"',8,'"&rs("Checker")&"','"&rs("CheckerID")&"','"&rs("CheckDate")&"','','',null,'"&rs("VPReplyer")&"','"&rs("VPReplyerID")&"','"&rs("VPReplyDate")&"','"&rs("Hrer")&"','"&rs("HrerID")&"','"&rs("HrDate")&"',4,0,'"&rs("Biller")&"','"&rs("BillerID")&"','"&rs("BillDate")&"',0,1,0,"&OSerialNum&"")
							totalhour=8
							if cdbl(left(etime,2))<13 then totalhour=4
							connzxpt.Execute("insert into Bill_OvertimeDetails select "&OSerialNum&",'"&rs2("Accompany")&"','"&rs2("AccompanyName")&"','"&rs2("AcDepartment")&"','"&rs2("AcDepartmentname")&"','"&rs2("Position")&"','"&dateadd("d",nnn,rs("StartDate"))&"','"&Mid(stime,1,6)&"','"&dateadd("d",nnn,rs("StartDate"))&"','"&Mid(etime,1,6)&"',8,'"&rs2("Grade")&"',0,0,0,"&totalhour&",'否'")
						end if
					next
					rs2.movenext
				wend
				rs2.close
				set rs2=nothing
			else
				response.write ("你没有权限进行此操作或当前状态不允许此次操作！")
				response.end
			end if
			rs.movenext
		wend
		rs.update
		rs.close
		set rs=nothing 
		response.write "审核成功！"
  elseif detailType="unCheck" then
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select * from Attendance_Travel where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
			if rs("CheckFlag")=0 or rs("CheckFlag")=3 then
				response.write ("此单据不允许反审核！")
				response.end
			end if
			if rs("CancelFlag") then
				response.Write("此单据已作废，不允许审核！")
				response.End()
			end if
'			if rs("CheckFlag")=3 and Instr(session("AdminPurviewFLW"),"|701.3,")>0 then
'				rs("Hrer")=AdminName
'				rs("HrerID")=UserName
'				rs("HrDate")=now()
'				rs("CheckFlag")=2
'			else
			if rs("CheckFlag")=2 and rs("VPReplyerID")=UserName then
				rs("VPReplyer")=AdminName
				rs("VPReplyerID")=UserName
				rs("VPReplyDate")=now()
				rs("CheckFlag")=2
			elseif rs("CheckFlag")=1 and rs("CheckerID")=UserName then
				rs("Checker")=AdminName
				rs("CheckerID")=UserName
				rs("CheckDate")=now()
				rs("CheckFlag")=0
			else
				response.write ("你没有权限进行此操作或当前状态不允许此次操作！")
				response.end
			end if
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
  if detailType="Register" or detailType="AgentID" then
    InfoID=request("InfoID")
		sql="select top 1 a.*,b.部门名称,case when 职等='001' then 3.75 else (FPA1004+FPA1005)/26/8 end as 时薪 from [N-基本资料单头] a,[G-部门资料表] b,AIS20081217153921.dbo.t_PANewData c,AIS20081217153921.dbo.t_PA_item d where a.部门别=b.部门代号 and a.员工代号 like '%"&InfoID&"%' and c.FEmpID=d.FItemID and FNumber=a.员工代号 order by FYear desc,FPeriod desc"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
				response.write ("员工编号不存在！")
				response.end
		else
			response.write(rs("员工代号")&"###"&rs("姓名")&"###"&rs("部门别")&"###"&rs("部门名称")&"###"&rs("性别")&"###"&rs("工作岗位")&"###"&rs("到职日")&"###"&rs("职等")&"###"&rs("时薪"))
		end if
		rs.close
		set rs=nothing 
  elseif detailType="SerialNum" then
    InfoID=request("InfoID")
		sql="select * from Attendance_Travel where SerialNum="&InfoID
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
			sql="select * from Attendance_TravelDetails where SNum="&InfoID
			set rs=server.createobject("adodb.recordset")
			rs.open sql,connzxpt,1,1
			do until rs.eof
				response.write "{"
				for i=0 to rs.fields.count-1
				if IsNull(rs.fields(i).value) then
				response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
				else
				response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
				end if
				next
				response.write ("""bg"":""#EBF2F9""}")
				rs.movenext
			If Not rs.eof Then
				Response.Write ","
			End If
			loop
			response.write "]}]}"
		end if
		rs.close
		set rs=nothing 
	end if
end if
 %>
