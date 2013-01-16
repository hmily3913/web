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
      datafrom=" Bill_Overtime a left join Bill_OvertimeDetails b on a.SerialNum=b.SNum "
  dim datawhere'数据条件
	if Instr(session("AdminPurviewFLW"),"|214.6,")>0 or Instr(session("AdminPurviewFLW"),"|214.5,")>0 then
		datawhere=" where left(Department,9)='"&left(Depart,9)&"' "
	elseif Instr(session("AdminPurviewFLW"),"|214.3,")>0 and Depart="KD01.0001.0003" then
    datawhere=" where (Department='"&Depart&"' or BillerID='"&UserName&"' or OverType='收料加班') "
	else
		datawhere=" where left(Department,9)='"&left(Depart,9)&"' "
'    datawhere=" where (Department='"&Depart&"' or BillerID='"&UserName&"') "
	end if
	Dim searchterm,searchcols
	if Request.Form("query") <> "" then
	searchterm = Request.Form("query")
	searchcols = request.Form("qtype")
		if Request.Form("qtype")="CtrlDate" then
			datawhere = datawhere&" and datediff(d,"&searchcols&",'"&searchterm&"')=0 "
		else
			datawhere = datawhere&" and "&searchcols&" like '%"&searchterm&"%' "
		end if
	else
	datawhere = datawhere&" and a.CancelFlag=0 "
	End if
	datawhere = datawhere&Session("AllMessage26")&Session("AllMessage27")&Session("AllMessage28")&Session("AllMessage29")&Session("AllMessage30")
	session.contents.remove "AllMessage26"
	session.contents.remove "AllMessage27"
	session.contents.remove "AllMessage28"
	session.contents.remove "AllMessage29"
	session.contents.remove "AllMessage30"
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
  sql="select count(distinct SerialNum) as idCount from "& datafrom &" " & datawhere
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
    sql="select distinct SerialNum,CheckFlag from "& datafrom &" " & datawhere & taxis
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
    sql="select distinct a.* from "& datafrom &" where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    do until rs.eof'填充数据到表格
	dim ys:ys="#f7f7f7"
	if rs("CheckFlag")>0 then ys="#ffff66"
	dim CheckState:CheckState=""
	if rs("CheckFlag")="1" then
	  CheckState="主管审核"
	elseif rs("CheckFlag")="2" then
	  CheckState="相关部门审核"
	elseif rs("CheckFlag")="3" then
	  CheckState="副总审核"
	elseif rs("CheckFlag")="4" then
	  CheckState="已执行"
	elseif rs("CheckFlag")="5" then
	  CheckState="已计算"
	else
	  CheckState="待审核"
	end if
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("Register")%>","<%=rs("RegisterName")%>","<%=rs("RegDate")%>","<%=rs("Departmentname")%>","<%=rs("AllHour")%>","<%=rs("OverType")%>","<%=JsonStr(rs("LeaveReason"))%>","<%=rs("ResponseDepart")%>","<%=rs("OversNum")%>","<%=rs("JCALL")%>","<%=CheckState%>","<%=rs("CancelFlag")%>","<%=rs("Biller")%>","<%=rs("BillDate")%>","<%=rs("Checker")%>","<%=rs("Relater")%>","<%=rs("VPReplyer")%>","<%=rs("Hrer")%>"]}
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
			SerialNum=getBillNo("Bill_Overtime",3,date())
			set rs = server.createobject("adodb.recordset")
			sql="select * from Bill_Overtime"
			rs.open sql,connzxpt,1,3
			rs.addnew
			rs("SerialNum")=SerialNum
			rs("Biller")=AdminName
			rs("BillerID")=UserName
			rs("BillDate")=now()
			rs("Register")=Request("Register")
			rs("RegisterName")=Request("RegisterName")
			rs("RegDate")=Request("RegDate")
			rs("Department")=Request("Department")
			rs("Departmentname")=Request("Departmentname")
			rs("OverType")=Request("OverType")
			rs("OverReason")=Request("OverReason")
			rs("ResponseDepart")=Request("ResponseDepart")
			dim allhour:allhour=0
			dim allfy:allfy=0
			dim OversNum:OversNum=0
			dim JCAll:JCAll=0
			for   i=2   to   Request.form("SerialNumD").count
				if (Request.form("DeleteFlag")(i)="0" or Request.form("DeleteFlag")(i)="") and Request.form("Overer")(i)<>"" then
					dim FYPrice
					if Request.form("Grade")(i)="001" then
						FYPrice=3.75
					else
						set rs2=connk3.Execute("select top 1 (FPA1004+FPA1005)/26/8 as sx from t_PANewData a,t_PA_item b where a.FEmpID=b.FItemID and b.FNumber='"&Request.form("Overer")(i)&"' order by FYear desc,FPeriod desc")
						if isnull(rs2("sx")) then
						FYPrice=0
						else
						FYPrice=rs2("sx")
						end if
					end if
					connzxpt.Execute("insert into Bill_OvertimeDetails (SNum,Overer,OvererName,OverDepartment,OverDepartmentname,Position,StartDate,StartTime,EndDate,EndTime,TotalHour,Grade,FYPrice,FYTotal,ActualHour,Repast) values ('"&SerialNum&"','"&Request.form("Overer")(i)&"','"&Request.form("OvererName")(i)&"','"&Request.form("OverDepartment")(i)&"','"&Request.form("OverDepartmentname")(i)&"','"&Request.form("Position")(i)&"','"&Request.form("StartDate")(i)&"','"&Request.form("StartTime")(i)&"','"&Request.form("EndDate")(i)&"','"&Request.form("EndTime")(i)&"','"&Request.form("TotalHour")(i)&"','"&Request.form("Grade")(i)&"','"&FYPrice&"','"&(FYPrice*cdbl(Request.form("TotalHour")(i)))&"','"&Request.form("TotalHour")(i)&"','"&Request.form("Repast")(i)&"')")
					allhour=allhour+cdbl(Request.form("TotalHour")(i))
					allfy=allfy+FYPrice*cdbl(Request.form("TotalHour")(i))
					OversNum=OversNum+1
					if Request.form("Repast")(i)="是" then JCAll=JCAll+1
				end if
			next
			if allhour>0 then
				rs.update
				rs.close
				set rs=nothing 
				connzxpt.Execute("update Bill_Overtime set AllHour="&allhour&",FYAll="&allfy&",OversNum="&OversNum&",JCAll="&JCAll&" where serialnum="&SerialNum)
				response.write "保存成功！"
			else
				response.Write("加班时间为0，保存失败，请检查！")
				response.End()
			end if
'		else
'			response.write ("你没有权限进行此操作！")
'			response.end
'		end if
  elseif detailType="Edit"  then
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select * from Bill_Overtime where SerialNum="&SerialNum
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
		rs("OverType")=Request("OverType")
		rs("OverReason")=Request("OverReason")
		rs("ResponseDepart")=Request("ResponseDepart")
    rs.update
		allhour=0
		OversNum=0
		JCAll=0
		for   i=2   to   Request.form("SerialNumD").count
			if Request.form("DeleteFlag")(i)="1" and Request.Form("SerialNumD")(i)<>"" then
				connzxpt.Execute("Delete from Bill_OvertimeDetails where SerialNumD="&Request.Form("SerialNumD")(i))
			elseif Request.Form("SerialNumD")(i)<>"" then
					if Request.form("Grade")(i)="001" then
						FYPrice=3.75
					else
						set rs=connk3.Execute("select top 1 (FPA1004+FPA1005)/26/8 as sx from t_PANewData a,t_PA_item b where a.FEmpID=b.FItemID and b.FNumber='"&Request.form("Overer")(i)&"' order by FYear desc,FPeriod desc")
						if rs("sx")="" then
						FYPrice=0
						else
						FYPrice=rs("sx")
						end if
					end if
				connzxpt.Execute("update Bill_OvertimeDetails set Overer='"&Request.form("Overer")(i)&"',OvererName='"&Request.form("OvererName")(i)&"',OverDepartment='"&Request.form("OverDepartment")(i)&"',OverDepartmentname='"&Request.form("OverDepartmentname")(i)&"',Position='"&Request.form("Position")(i)&"',StartDate='"&Request.form("StartDate")(i)&"',StartTime='"&Request.form("StartTime")(i)&"',EndDate='"&Request.form("EndDate")(i)&"',EndTime='"&Request.form("EndTime")(i)&"',TotalHour='"&Request.form("TotalHour")(i)&"',Grade='"&Request.form("Grade")(i)&"',FYPrice='"&FYPrice&"',FYTotal="&(FYPrice*cdbl(Request.form("TotalHour")(i)))&",ActualHour='"&Request.form("TotalHour")(i)&"',Repast='"&Request.form("Repast")(i)&"' where SerialNumD="&Request.Form("SerialNumD")(i))
				allhour=allhour+cdbl(Request.form("TotalHour")(i))
				allfy=allfy+FYPrice*cdbl(Request.form("TotalHour")(i))
				OversNum=OversNum+1
				if Request.form("Repast")(i)="是" then JCAll=JCAll+1
			elseif Request.form("Overer")(i)<>"" then
					if Request.form("Grade")(i)="001" then
						FYPrice=3.75
					else
						set rs=connk3.Execute("select top 1 (FPA1004+FPA1005)/26/8 as sx from t_PANewData a,t_PA_item b where a.FEmpID=b.FItemID and b.FNumber='"&Request.form("Overer")(i)&"' order by FYear desc,FPeriod desc")
						if isnull(rs("sx")) then
						FYPrice=0
						else
						FYPrice=rs("sx")
						end if
					end if
					connzxpt.Execute("insert into Bill_OvertimeDetails (SNum,Overer,OvererName,OverDepartment,OverDepartmentname,Position,StartDate,StartTime,EndDate,EndTime,TotalHour,Grade,FYPrice,FYTotal,ActualHour,Repast) values ('"&SerialNum&"','"&Request.form("Overer")(i)&"','"&Request.form("OvererName")(i)&"','"&Request.form("OverDepartment")(i)&"','"&Request.form("OverDepartmentname")(i)&"','"&Request.form("Position")(i)&"','"&Request.form("StartDate")(i)&"','"&Request.form("StartTime")(i)&"','"&Request.form("EndDate")(i)&"','"&Request.form("EndTime")(i)&"','"&Request.form("TotalHour")(i)&"','"&Request.form("Grade")(i)&"','"&FYPrice&"','"&(FYPrice*cdbl(Request.form("TotalHour")(i)))&"','"&Request.form("TotalHour")(i)&"','"&Request.form("Repast")(i)&"')")
					allhour=allhour+cdbl(Request.form("TotalHour")(i))
					allfy=allfy+FYPrice*cdbl(Request.form("TotalHour")(i))
					OversNum=OversNum+1
					if Request.form("Repast")(i)="是" then JCAll=JCAll+1
			end if
		next
		connzxpt.Execute("update Bill_Overtime set AllHour="&allhour&",FYAll="&allfy&",OversNum="&OversNum&",JCAll="&JCAll&" where serialnum="&SerialNum)
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
		sql="select * from Bill_Overtime where CheckFlag>0 and SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,1
		if rs.bof and rs.eof then
			connzxpt.Execute("Delete from Bill_Overtime where SerialNum in ("&SerialNum&")")
			connzxpt.Execute("Delete from Bill_OvertimeDetails where SNum in ("&SerialNum&")")
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
		sql="select * from Bill_Overtime where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
			if rs("CancelFlag") then
				response.Write("此单据已作废，不允许审核！")
				response.End()
			end if
			if rs("CheckFlag")=2 and Instr(session("AdminPurviewFLW"),"|214.5,")>0 then
				rs("VPReplyer")=AdminName
				rs("VPReplyerID")=UserName
				rs("VPReplyDate")=now()
				rs("CheckFlag")=3
			elseif rs("CheckFlag")=1 then
				if rs("OverType")<>"收料加班" and Instr(session("AdminPurviewFLW"),"|214.5,")>0 then
					rs("VPReplyer")=AdminName
					rs("VPReplyerID")=UserName
					rs("VPReplyDate")=now()
					rs("CheckFlag")=3
				end if
'				if (rs("Department")="KD01.0005.0004" or rs("OverType")="接送客人") and Instr(session("AdminPurviewFLW"),"|214.4,")>0 then
'					rs("VPReplyer")=AdminName
'					rs("VPReplyerID")=UserName
'					rs("VPReplyDate")=now()
'					rs("CheckFlag")=3
'				end if
'				if (rs("OverType")="正常加班" or rs("OverType")="成品出货") and rs("Department")<>"KD01.0005.0004" and Instr(session("AdminPurviewFLW"),"|214.5,")>0 then
'					rs("VPReplyer")=AdminName
'					rs("VPReplyerID")=UserName
'					rs("VPReplyDate")=now()
'					rs("CheckFlag")=3
'				end if
'				if (rs("OverType")="分厂加班" or rs("OverType")="外协加班") and Instr(session("AdminPurviewFLW"),"|214.3,")>0 and Depart="KD01.0001.0017" then
'					rs("Relater")=AdminName
'					rs("RelaterID")=UserName
'					rs("RelatDate")=now()
'					rs("CheckFlag")=2
'				end if
				if rs("OverType")="收料加班" and Instr(session("AdminPurviewFLW"),"|214.3,")>0 and Depart="KD01.0001.0003" then
					rs("Relater")=AdminName
					rs("RelaterID")=UserName
					rs("RelatDate")=now()
					rs("CheckFlag")=2
				end if
			elseif rs("CheckFlag")=0 and Instr(session("AdminPurviewFLW"),"|214.2,")>0 then
				if rs("OversNum")=0 then
					response.Write("加班人数为0，不允许审核，请检查！")
					response.End()
				end if
				if Depart=rs("Department") or (rs("Department")="KD01.0001.0005" and Depart="KD01.0001.0012") then
					rs("Checker")=AdminName
					rs("CheckerID")=UserName
					rs("CheckDate")=now()
					rs("CheckFlag")=1
				else
					response.Write("只能审核本部门的加班申请单！")
					response.End()
				end if
			else
				response.write ("你没有权限进行此操作或当前状态不允许此次操作！")
				response.end
			end if
			rs("CancelFlag")=request("operattext")
			rs.movenext
		wend
		rs.update
		rs.close
		set rs=nothing 
		response.write "审核成功！"
  elseif detailType="Count" then
		if Instr(session("AdminPurviewFLW"),"|214.7,")=0 then
			response.Write("您没有权限进行此操作，请检查！")
			response.End()
		end if
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select * from Bill_Overtime where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
			if rs("CheckFlag")=4 then
				rs("CheckFlag")=5
			else
				response.Write("只能计算已实施的单据，请检查！")
				response.End()
			end if
			rs.movenext
		wend
		rs.update
		rs.close
		set rs=nothing 
		response.write "计算成功！"
  elseif detailType="Finish" then
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select * from Bill_Overtime where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs("CancelFlag") then
			response.Write("此单据已作废，不允许审核！")
			response.End()
		end if
		if rs("CheckFlag")=3 and Instr(session("AdminPurviewFLW"),"|214.6,")>0 then
			rs("Hrer")=AdminName
			rs("HrerID")=UserName
			rs("HrDate")=now()
			rs("CheckFlag")=4
			rs.update
			rs.close
			set rs=nothing 
		for   i=2   to   Request.form("SerialNumD").count
			if (Request.form("DeleteFlag")(i)="0" or Request.form("DeleteFlag")(i)="") and Request.Form("SerialNumD")(i)<>"" then
				if Request.form("Grade")(i)="001" then
					FYPrice=3.75
				else
					set rs=connk3.Execute("select top 1 (FPA1004+FPA1005)/26/8 as sx from t_PANewData a,t_PA_item b where a.FEmpID=b.FItemID and b.FNumber='"&Request.form("Overer")(i)&"' order by FYear desc,FPeriod desc")
					if rs("sx")="" then
					FYPrice=0
					else
					FYPrice=rs("sx")
					end if
				end if
				connzxpt.Execute("update Bill_OvertimeDetails set FYPrice='"&FYPrice&"',FYTotal="&(FYPrice*cdbl(Request.form("ActualHour")(i)))&",ActualHour='"&Request.form("ActualHour")(i)&"' where SerialNumD="&Request.Form("SerialNumD")(i))
			end if
		next
		connzxpt.Execute("update Bill_Overtime set AllHour=aaa.c,FYAll=aaa.b,OversNum=aaa.a from (select count(1) a,sum(FYTotal) b,sum(ActualHour) c,SNum from Bill_OvertimeDetails group by SNum) aaa where aaa.SNum=SerialNum and SerialNum="&SerialNum)
		response.write "执行成功"
		else
			response.Write("该单据不允许此操作或者你没有权限进行此操作！")
		end if
  elseif detailType="unCheck" then
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select * from Bill_Overtime where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
			if rs("CheckFlag")=0 then
				response.write ("此单据未审核，不允许反审核！")
				response.end
			end if
			if rs("CancelFlag") then
				response.Write("此单据已作废，不允许审核！")
				response.End()
			end if
			if rs("CheckFlag")=5 and Instr(session("AdminPurviewFLW"),"|214.7,")>0 then
				rs("CheckFlag")=4
			elseif rs("CheckFlag")=4 and Instr(session("AdminPurviewFLW"),"|214.6,")>0 then
				rs("Hrer")=AdminName
				rs("HrerID")=UserName
				rs("HrDate")=now()
				rs("CheckFlag")=3
			elseif rs("CheckFlag")=3 and rs("VPReplyerID")=UserName then
				rs("VPReplyer")=AdminName
				rs("VPReplyerID")=UserName
				rs("VPReplyDate")=now()
				if rs("RelaterID")="" then
				rs("CheckFlag")=1
				else
				rs("CheckFlag")=2
				end if
			elseif rs("CheckFlag")=2 and rs("RelaterID")=UserName then
				rs("Relater")=AdminName
				rs("RelaterID")=UserName
				rs("RelatDate")=now()
				rs("CheckFlag")=1
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
  if detailType="Register" then
    InfoID=request("InfoID")
		sql="select top 1 a.*,b.部门名称,case when 职等='001' then 3.75 else isnull((FPA1004+FPA1005)/26/8,0) end as 时薪 from [N-基本资料单头] a inner join [G-部门资料表] b on a.部门别=b.部门代号 inner join AIS20081217153921.dbo.t_PA_item d on FNumber=a.员工代号 left join AIS20081217153921.dbo.t_PANewData c on c.FEmpID=d.FItemID where a.员工代号 like '%"&InfoID&"%'  order by FYear desc,FPeriod desc"
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
		sql="select * from Bill_Overtime where SerialNum="&InfoID
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
			sql="select * from Bill_OvertimeDetails where SNum="&InfoID
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
				sqlkq="select min(checktime) as mint,max(checktime) aS maxt from USERINFO a,CHECKINOUT b where a.userid=b.userid and a.ssn='"&rs("overer")&"' and datediff(d,checktime,'"&rs("enddate")&"')>=0 and datediff(d,checktime,'"&rs("startdate")&"')<=0  "
				set rskq=server.createobject("adodb.recordset")
				rskq.open sqlkq,connkq,1,1
				if rskq.eof or isnull(rskq("mint")) or datediff("s",rskq("mint"),rs("startdate")&" "&rs("starttime")&":00")<0 or datediff("s",rskq("maxt"),rs("enddate")&" "&rs("endtime")&":00")>0 then 
				response.write ("""bg"":""#ff99ff""}")
				else
				response.write ("""bg"":""#EBF2F9""}")
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
  elseif detailType="Users" then
		dim tempstr1
    InfoID=request("SerialNum")
		if InfoID="" then
			tempstr1=" and pspk='0'"
		else
			tempstr1=" and pspk='"&InfoID&"'"
		end if
    set rs = server.createobject("adodb.recordset")
    sql="select * from (select a.员工代号 as pk,a.员工代号+'/'+a.姓名 as name,c.部门代号 as pspk,c.部门名称 as val1,a.工作岗位 as val2,a.职等 as val3,a.姓名 as val4 from [N-基本资料单头] a,[G-部门资料表] c  where a.离职否='在职' and a.二级部门=c.部门代号  "
    sql=sql&" union all "
    sql=sql&" select distinct c.部门代号 as pk,c.部门名称 as name,'0' as pspk,'' as val1,'' as val2,'' as val3,'' as val4 from [N-基本资料单头] a,[G-部门资料表] c  where a.离职否='在职' and a.部门别=c.部门代号"
    sql=sql&" union all "
    sql=sql&" select distinct a.二级部门 as pk,c.部门名称 as name,a.部门别 as pspk,'' as val1,'' as val2,'' as val3,'' as val4 from [N-基本资料单头] a,[G-部门资料表] c  where a.离职否='在职' and a.二级部门=c.部门代号 and a.二级部门<>a.部门别) aaa where 1=1"&tempstr1
    rs.open sql,conn,1,1
		response.Write("[")
		do until rs.eof
	%>
	{"SerialNum": "<%=rs("pk")%>", "name": "<%=rs("name")%>", "PSNum": "<%=rs("pspk")%>","val1":"<%=rs("val1")%>","val2":"<%=rs("val2")%>","val3":"<%=rs("val3")%>","val4":"<%=rs("val4")%>"
	<%
		if rs("val1")="" then
	%>
	, isParent:true
	<%
		end if
		response.Write("}")
			rs.movenext
			If Not rs.eof Then
			  Response.Write ","
			End If
		loop
		Response.Write "]"
		rs.close
		set rs=nothing 
  end if
elseif showType="xls2sql" then
	Server.ScriptTimeout = 999999
	set rs=server.createobject("adodb.recordset")
	sql="select * from Bill_Overtime"
	rs.open sql,connzxpt,1,3
	InfoID=request("InfoID")
	Set xlApp=Server.CreateObject("Excel.Application")          '/******** VBA方法 连接Excel *********/
	Set xlbook=xlApp.Workbooks.Open(Server.mappath(InfoID))  
	Set xlsheet=xlbook.Worksheets(1)  
	i=2
	While cstr(xlsheet.cells(i,1))<>""           '/********** 使用第3列 帐号为空时判断为结束标志  **********/
		rs.Addnew()  
		rs("Biller")=AdminName
		rs("BillerID")=UserName
		rs("BillDate")=now()
		rs("Register")=xlsheet.cells(i,1)
		rs("RegisterName")=xlsheet.cells(i,2)
		rs("RegDate")=xlsheet.cells(i,6)
		rs("Department")=xlsheet.cells(i,4)
		rs("Departmentname")=xlsheet.cells(i,5)
		rs("StartDate")=xlsheet.cells(i,8)
		rs("StartTime")=xlsheet.cells(i,9)
		rs("EndDate")=xlsheet.cells(i,10)
		rs("EndTime")=xlsheet.cells(i,11)
		rs("TotalHour")=xlsheet.cells(i,12)
		rs("Position")=xlsheet.cells(i,3)
		rs("OverType")=xlsheet.cells(i,7)
		rs("OverReason")=xlsheet.cells(i,13)
	i=i+1  
	Wend  
	rs.Update  
	rs.Close
	Set rs=Nothing
	xlsheet.close
	Set xlsheet=nothing  
	xlbook.Close  
	Set xlbook=Nothing  
	xlApp.DisplayAlerts=false
	xlApp.Quit  
	response.Write("共计"&i-2&"条数据导入成功!")
elseif showType="CheckPermiss" then
	if Instr(session("AdminPurviewFLW"),"|214.5,")>0 or Instr(session("AdminPurviewFLW"),"|214.4,")>0 then
		response.Write("True")
	else
		response.Write("False")
	end if
elseif showType="Excel" then 
	sql="exec sp_overtimes "&request("Year")&","&request("Month")&",'"&request("PesonType")&"' "
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
elseif showType="OutPutOne" then 
	InfoID=request("InfoID")
	response.ContentType("application/vnd.ms-excel")
	response.AddHeader "Content-disposition", "attachment; filename=erpData.xls"
	sql="select * from Bill_OvertimeDetails where SNum="&InfoID
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
%>
		<table width="100%" border="0" id="editDetails" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
		  <tbody id="TbDetails">
		  <tr height="24">
			<td ><font color="#FFFFFF"><strong>加班人工号</strong></font></td>
			<td ><font color="#FFFFFF"><strong>姓名</strong></font></td>
			<td ><font color="#FFFFFF"><strong>部门编号</strong></font></td>
			<td ><font color="#FFFFFF"><strong>部门名称</strong></font></td>
			<td ><font color="#FFFFFF"><strong>岗位</strong></font></td>
			<td ><font color="#FFFFFF"><strong>职等</strong></font></td>
			<td ><font color="#FFFFFF"><strong>开始日期</strong></font></td>
			<td ><font color="#FFFFFF"><strong>开始时间</strong></font></td>
			<td ><font color="#FFFFFF"><strong>结束日期</strong></font></td>
			<td ><font color="#FFFFFF"><strong>结束时间</strong></font></td>
			<td ><font color="#FFFFFF"><strong>总计/小时</strong></font></td>
			<td ><font color="#FFFFFF"><strong>实际/小时</strong></font></td>
			<td ><font color="#FFFFFF"><strong>是否就餐</strong></font></td>
		  </tr>
<%
	do until rs.eof
%>
	<tr height="24" bgcolor="#EBF2F9">
			<td ><%=rs("Overer")%></td>
			<td ><%=rs("OvererName")%></td>
			<td ><%=rs("OverDepartment")%></td>
			<td ><%=rs("OverDepartmentname")%></td>
			<td ><%=rs("Position")%></td>
			<td ><%=rs("Grade")%></td>
			<td ><%=rs("StartDate")%></td>
			<td ><%=rs("StartTime")%></td>
			<td ><%=rs("EndDate")%></td>
			<td ><%=rs("EndTime")%></td>
			<td ><%=rs("TotalHour")%></td>
			<td ><%=rs("ActualHour")%></td>
			<td ><%=rs("Repast")%></td>
   </tr>
<%
		rs.movenext
	loop
	response.Write("</table>")
	rs.close
	set rs=nothing 
elseif showType="showkq" then 
	sqlkq="select checktime from USERINFO a,CHECKINOUT b where a.userid=b.userid and a.ssn='"&request("ssn")&"' and datediff(d,checktime,'"&request("date")&"')=0 "
	set rskq=server.createobject("adodb.recordset")
	rskq.open sqlkq,connkq,1,1
	if rskq.eof then
		response.Write("无考勤")
	else
		while (not rskq.eof)
			response.Write("【"&right(rskq("checktime"),8)&"】")
			rskq.movenext()
		wend
	end if
elseif showType="JCOutPut" then 
	sql="select * from Bill_Overtime a ,Bill_OvertimeDetails b where a.SerialNum=b.SNum and b.Repast='是' "
	if request("SDate")<>"" then sql=sql&" and StartDate>='"&request("SDate")&"' "
	if request("EDate")<>"" then sql=sql&" and EndDate<='"&request("EDate")&"' "
	sql=sql&" order by OverDepartment,Overer,StartDate "
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
%>
		<table width="100%" border="0" id="editDetails" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
		  <tbody id="TbDetails">
		  <tr height="24">
			<td ><font color="#FFFFFF"><strong>部门编号</strong></font></td>
			<td ><font color="#FFFFFF"><strong>部门名称</strong></font></td>
			<td ><font color="#FFFFFF"><strong>加班人工号</strong></font></td>
			<td ><font color="#FFFFFF"><strong>姓名</strong></font></td>
			<td ><font color="#FFFFFF"><strong>岗位</strong></font></td>
			<td ><font color="#FFFFFF"><strong>职等</strong></font></td>
			<td ><font color="#FFFFFF"><strong>开始日期</strong></font></td>
			<td ><font color="#FFFFFF"><strong>开始时间</strong></font></td>
			<td ><font color="#FFFFFF"><strong>结束日期</strong></font></td>
			<td ><font color="#FFFFFF"><strong>结束时间</strong></font></td>
		  </tr>
<%	
	do until rs.eof
%>
	<tr height="24" bgcolor="#EBF2F9">
			<td ><%=rs("OverDepartment")%></td>
			<td ><%=rs("OverDepartmentname")%></td>
			<td ><%=rs("Overer")%></td>
			<td ><%=rs("OvererName")%></td>
			<td ><%=rs("Position")%></td>
			<td ><%=rs("Grade")%></td>
			<td ><%=rs("StartDate")%></td>
			<td ><%=rs("StartTime")%></td>
			<td ><%=rs("EndDate")%></td>
			<td ><%=rs("EndTime")%></td>
   </tr>
<%
		rs.movenext
	loop
	response.Write("</table>")
	rs.close
	set rs=nothing 
end if
 %>
