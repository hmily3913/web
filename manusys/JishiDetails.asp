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
DepartName=session("DepartName")
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
      datafrom=" manusys_Jishi a left join manusys_JishiDetails1 b on a.SerialNum=b.SNum "
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
	else
	End if
		 if request("sd")<>"" then datawhere=datawhere&" and datediff(d,a.RegDate,'"&request("sd")&"')<=0 "
		 if request("ed")<>"" then datawhere=datawhere&" and datediff(d,a.RegDate,'"&request("ed")&"')>=0 "
		 if request("ssd")<>"" then datawhere=datawhere&" and datediff(d,a.CheckDate,'"&request("ssd")&"')<=0 "
		 if request("sed")<>"" then datawhere=datawhere&" and datediff(d,a.CheckDate,'"&request("sed")&"')>=0 "
		 if request("bm")<>"" then datawhere=datawhere&" and a.Bumen='"&request("bm")&"' "
		 if request("dd")<>"" then datawhere=datawhere&" and b.OrderID like '%"&request("dd")&"%' "
		 if request("sh")<>"" then datawhere=datawhere&" and a.CheckFlag='"&request("sh")&"' "
	datawhere = datawhere&Session("AllMessage72")&Session("AllMessage73")&Session("AllMessage74")&Session("AllMessage75")
	session.contents.remove "AllMessage72"
	session.contents.remove "AllMessage73"
	session.contents.remove "AllMessage74"
	session.contents.remove "AllMessage75"
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
    sql="select SerialNumD from "& datafrom &" " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    rs.pagesize = pages '每页显示记录数
	rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
	  if(i=1)then
	    sqlid=rs("SerialNumD")
	  else
	    sqlid=sqlid &","&rs("SerialNumD")
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
    sql="select a.*,b.* from "& datafrom &" where SerialNumD in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    do until rs.eof'填充数据到表格
	dim CheckState:CheckState=""
	if rs("CheckFlag")="1" then
	  CheckState="段组长审核"
	elseif rs("CheckFlag")="2" then
	  CheckState="厂长审核"
	elseif rs("CheckFlag")="3" then
	  CheckState="责任确认"
	elseif rs("CheckFlag")="4" then
	  CheckState="已执行"
	else
	  CheckState="待审核"
	end if
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("Bumen")%>","<%=rs("RegDate")%>","<%=rs("EmpName")%>","<%=rs("OrderID")%>","<%=rs("Product")%>","<%=JsonStr(rs("ReasonType"))%>","<%=JsonStr(rs("Reason"))%>","<%=rs("StartTime")%>","<%=rs("EndTime")%>","<%=rs("Shijian")%>","<%=rs("Wancheng")%>","<%=rs("Danjia")%>","<%=rs("Jine")%>","<%=CheckState%>","<%=rs("Biller")%>","<%=rs("BillDate")%>","<%=rs("Checker1")%>","<%=rs("Checker2")%>"]}
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
			SerialNum=getBillNo("manusys_Jishi",3,date())
			set rs = server.createobject("adodb.recordset")
			sql="select * from manusys_Jishi"
			rs.open sql,connzxpt,1,3
			rs.addnew
			rs("SerialNum")=SerialNum
			rs("Biller")=AdminName
			rs("BillerID")=UserName
			rs("BillDate")=now()
			rs("Bumen")=Request("Bumen")
			rs("RegDate")=Request("RegDate")
			rs("Remark")=Request("Remark")
			rs.update
		for   i=2   to   Request.form("SerialNumD").count
			if request.Form("EmpID")(i)<>"" and request.Form("Shijian")(i)<>"" and request.Form("Shijian")(i)<>"0" then
			set rs = server.createobject("adodb.recordset")
			sql="select * from manusys_JishiDetails1"
			rs.open sql,connzxpt,1,3
			rs.addnew
			rs("SNum")=SerialNum
			rs("EmpID")=request.Form("EmpID")(i)
			rs("EmpName")=request.Form("EmpName")(i)
			rs("OrderID")=request.Form("OrderID")(i)
			rs("Product")=request.Form("Product")(i)
			rs("ReasonType")=request.Form("ReasonType")(i)
			rs("Reason")=request.Form("Reason")(i)
			rs("StartTime")=request.Form("StartTime")(i)
			rs("EndTime")=request.Form("EndTime")(i)
			rs("Shijian")=request.Form("Shijian")(i)
			rs("Wancheng")=request.Form("Wancheng")(i)
			rs("Danjia")=request.Form("Danjia")(i)
			rs("Jine")=cdbl(request.Form("Danjia")(i))*cdbl(request.Form("Shijian")(i))
			rs.update
			end if
		next
		for   i=2   to   Request.form("SerialNumD2").count
			if request.Form("Zerenbume")(i)<>"" then
			set rs = server.createobject("adodb.recordset")
			sql="select * from manusys_JishiDetails2"
			rs.open sql,connzxpt,1,3
			rs.addnew
			rs("SNum")=SerialNum
			rs("Zerenbume")=request.Form("Zerenbume")(i)
			rs("Shuoming")=request.Form("Shuoming")(i)
			rs("Checker")=request.Form("Checker")(i)
			rs("CheckerID")=request.Form("CheckerID")(i)
			rs.update
			end if
		next
		rs.close
		set rs=nothing 
		response.write "保存成功！"
'		else
'			response.write ("你没有权限进行此操作！")
'			response.end
'		end if
  elseif detailType="Edit"  then
			SerialNum=request("SerialNum")
			set rs = server.createobject("adodb.recordset")
			sql="select * from manusys_Jishi where SerialNum="&Request("SerialNum")
			rs.open sql,connzxpt,1,3
			if rs("CheckFlag")>1 then
				response.Write("厂长已审核不允许修改！")
				response.End()
			end if
			if (DepartName<>rs("Bumen") or (Instr(session("AdminPurview"),"|311.3,")=0 and Instr(session("AdminPurview"),"|311.2,")=0)) and rs("BillerID")<>UserName then
				response.Write("你没有权限修改此单据！")
				response.End()
			end if
			rs("Biller")=AdminName
			rs("BillerID")=UserName
			rs("BillDate")=now()
			rs("Bumen")=Request("Bumen")
			rs("RegDate")=Request("RegDate")
			rs("Remark")=Request("Remark")
			rs.update
		for   i=2   to   Request.form("SerialNumD").count
			if request.Form("EmpID")(i)<>"" and request.Form("Shijian")(i)<>"" then
			set rs = server.createobject("adodb.recordset")
			sql="select * from manusys_JishiDetails1 "
			if Request.form("SerialNumD")(i)<>"" then sql=sql&" where SerialNumD="&Request.form("SerialNumD")(i)
			rs.open sql,connzxpt,1,3
			if Request.form("SerialNumD")(i)="" then rs.addnew
			rs("SNum")=SerialNum
			rs("EmpID")=request.Form("EmpID")(i)
			rs("EmpName")=request.Form("EmpName")(i)
			rs("OrderID")=request.Form("OrderID")(i)
			rs("Product")=request.Form("Product")(i)
			rs("ReasonType")=request.Form("ReasonType")(i)
			rs("Reason")=request.Form("Reason")(i)
			rs("StartTime")=request.Form("StartTime")(i)
			rs("EndTime")=request.Form("EndTime")(i)
			rs("Shijian")=request.Form("Shijian")(i)
			rs("Wancheng")=request.Form("Wancheng")(i)
			rs("Danjia")=request.Form("Danjia")(i)
			rs("Jine")=cdbl(request.Form("Danjia")(i))*cdbl(request.Form("Shijian")(i))
			rs.update
			end if
		next
		for   i=2   to   Request.form("SerialNumD2").count
			if request.Form("Zerenbume")(i)<>"" then
			set rs = server.createobject("adodb.recordset")
			sql="select * from manusys_JishiDetails2"
			if Request.form("SerialNumD2")(i)<>"" then sql=sql&" where SerialNumD2="&Request.form("SerialNumD2")(i)
			rs.open sql,connzxpt,1,3
			if Request.form("SerialNumD2")(i)="" then rs.addnew
			rs("SNum")=SerialNum
			rs("Zerenbume")=request.Form("Zerenbume")(i)
			rs("Shuoming")=request.Form("Shuoming")(i)
			rs("Checker")=request.Form("Checker")(i)
			rs("CheckerID")=request.Form("CheckerID")(i)
			rs.update
			end if
		next
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
		sql="select * from manusys_Jishi where CheckFlag>0 and SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,1
		if rs.bof and rs.eof then
			connzxpt.Execute("Delete from manusys_Jishi where SerialNum in ("&SerialNum&")")
			connzxpt.Execute("Delete from manusys_JishiDetails1 where SNum in ("&SerialNum&")")
			connzxpt.Execute("Delete from manusys_JishiDetails2 where SNum in ("&SerialNum&")")
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
		sql="select * from manusys_Jishi where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
			if rs("CheckFlag")=0 and Instr(session("AdminPurview"),"|311.2,")>0 then
				if DepartName=rs("Bumen") then
					rs("Checker1")=AdminName
					rs("CheckerID1")=UserName
					rs("CheckDate1")=now()
					rs("CheckFlag")=1
				else
					response.Write("只能审核本部门的加班申请单！")
					response.End()
				end if
			elseif rs("CheckFlag")=1 and Instr(session("AdminPurview"),"|311.3,")>0 then
				set rs2=connzxpt.Execute("select 1 from manusys_JishiDetails2 where SNum="&rs("SerialNum"))
				if rs2.eof then
					response.Write("责任部门不能为空，审核失败，请检查！")
					response.End()
				elseif DepartName=rs("Bumen") then
					rs("Checker2")=AdminName
					rs("CheckerID2")=UserName
					rs("CheckDate2")=now()
'					connzxpt.Execute("update manusys_JishiDetails2 set CheckFlag=1,Tongyi=1,Checker='"&AdminName&"',CheckerID='"&UserName&"',CheckDate='"&now()&"' where (Zerenbume='金乡' or Zerenbume='营销部') SNum="&rs("SerialNum"))
					set rs2=connzxpt.Execute("select 1 from manusys_JishiDetails2 where (CheckFlag=0 or Butongyi=1) and SNum="&rs("SerialNum"))
					if rs2.eof then
						rs("CheckFlag")=3
					else
						rs("CheckFlag")=2
					end if
				else
					response.Write("只能审核本部门的加班申请单！")
					response.End()
				end if
			elseif rs("CheckFlag")=2 and Instr(session("AdminPurview"),"|311.5,")>0 then
					rs("Checker")=AdminName
					rs("CheckerID")=UserName
					rs("CheckDate")=now()
					rs("Checker3")=AdminName
					rs("CheckerID3")=UserName
					rs("CheckDate3")=now()
					rs("CheckFlag")=3
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
		sql="select * from manusys_Jishi where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
			if rs("CheckFlag")=0 then
				response.write ("此单据未审核，不允许反审核！")
				response.end
			end if
			if rs("CheckFlag")=2 and rs("CheckerID2")=UserName and Instr(session("AdminPurview"),"|311.3,")>0 then
				set rs2=connzxpt.Execute("select 1 from manusys_JishiDetails2 where CheckFlag=1 and SNum="&rs("SerialNum"))
				if not rs2.eof then
					response.Write("责任部门已经会签，不允许反审核！")
					response.End()
				end if
				rs("Checker2")=null
				rs("CheckerID2")=null
				rs("CheckDate2")=null
				rs("CheckFlag")=1
			elseif rs("CheckFlag")=1 and rs("CheckerID1")=UserName and Instr(session("AdminPurview"),"|311.2,")>0 then
				rs("Checker1")=null
				rs("CheckerID1")=null
				rs("CheckDate1")=null
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
  elseif detailType="DeleteDetails" then
		SerialNum=request("SerialNum")
		sql="select CheckFlag,BillerID from manusys_Jishi where SerialNum ="&request("No")
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		If rs("BillerID")<>UserName then
			response.Write("只能删除自己的单据，请检查！")
			response.End()
		end if
		If rs("CheckFlag")>0 then
			response.Write("已审核不允许删除，请检查！")
			response.End()
		end if
		if request("Type")="t1" then
		connzxpt.Execute("Delete from manusys_JishiDetails1 where SerialNumD ="&SerialNum)
		response.write "删除成功！"
		elseif request("Type")="t2" then
		connzxpt.Execute("Delete from manusys_JishiDetails2 where SerialNumD2="&SerialNum)
		response.write "删除成功！"
		end if
  elseif detailType="CheckDetails" then
'		if Instr(session("AdminPurview"),"|311.3,")=0 then
'			response.Write("你没有权限进行此操作！")
'			response.End()
'		end if
		SerialNum=request("SerialNum")
		if request("Type")="Agree" or request("Type")="Disagree" then
			set rs=connzxpt.Execute("select CheckFlag from manusys_Jishi where SerialNum ="&request("No"))
			If rs("CheckFlag")<2 then
				response.Write("厂长未审核，不需要确认,请检查！")
				response.End()
			end if
			sql="select * from manusys_JishiDetails2 where SerialNumD2 ="&SerialNum
			set rs=server.createobject("adodb.recordset")
			rs.open sql,connzxpt,1,3
			if rs("CheckFlag")=1 then
				response.Write("已确认不允许重复操作，请检查！")
				response.End()
			end if
			if rs("CheckerID")<>UserName then
				response.Write("你没有权限进行此确认，请检查！")
				response.End()
			end if
			rs("CheckFlag")=1
			rs("CheckDate")=now()
			rs("Shuoming")=request("CheckText")
			if request("Type")="Agree" then
				rs("Tongyi")=1
			else
				rs("Butongyi")=1
			end if
			rs.update
			set rs2=connzxpt.Execute("select 1 from manusys_JishiDetails2 where (CheckFlag=0 or Butongyi=1) and SNum ="&request("No"))
			if rs2.eof then
				connzxpt.Execute("Update manusys_Jishi set CheckFlag=3,CheckDate='"&now()&"',CheckerID='"&UserName&"',Checker='"&AdminName&"' where SerialNum="&request("No"))
			end if
		else
			sql="select * from manusys_JishiDetails2 where SerialNumD2 ="&SerialNum
			set rs=server.createobject("adodb.recordset")
			rs.open sql,connzxpt,1,3
			if rs("CheckFlag")=0 then
				response.Write("未确认，不需要取消确认！")
				response.End()
			end if
			if rs("CheckerID")<>UserName then
				response.Write("只能取消自己确认的单据，请检查！")
				response.End()
			end if
			rs("CheckFlag")=0
			rs("CheckDate")=null
			rs("Shuoming")=null
			rs.update
			connzxpt.Execute("Update manusys_Jishi set CheckFlag=2,CheckDate=null,CheckerID=null,Checker=null where CheckFlag=3 and SerialNum="&request("No"))
		end if
		response.Write("审核成功！")
  elseif detailType="Finish" then
		if Instr(session("AdminPurview"),"|311.6,")=0 then
			response.Write("你没有权限进行此操作！")
			response.End()
		end if
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select a.*,b.* from manusys_Jishi a left join manusys_JishiDetails1 b on a.SerialNum=b.SNum where a.CheckFlag=3 and a.Bumen='"&DepartName&"' and a.SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,1
		'如果有需要执行的明细
		if not rs.eof then
			dim FInterID
			dim FBillNO
			dim k3userid
			dim FEntryID
			'获取k3用户ID
			set rs2=connk3.Execute("select FUserid from t_user where FName='"&AdminName&"'")
			if rs2.eof then
				response.Write("k3中无权限，不能执行！")
				response.End()
			else
				k3userid=rs2("FUserid")
			end if
			'取FInterID开始----------------------------
			'旧的方式
'			set rs2=connk3.Execute("select isnull(FMaxNum,0) as FInterID from ICMaxNum where FTableName= 'ICJobPay'")
'			FInterID=rs2("FInterID")
'			connk3.Execute("update ICMaxNum set FMaxNum="&FInterID+1&" where FTableName= 'ICJobPay'")
			'修改后新的方式,直接调用k3原来的存储过程.
			Dim adoComm,prm
			'// 创建一个对象，我们用来调用存储过程 
			Set adoComm = CreateObject("ADODB.Command") 
			With adoComm 
			'// 设置连接，假设 adoConn 为已经连接的 ADODB.Connection 对象 
			.ActiveConnection = connk3 
			'// 类型为存储过程，adCmdStoredProc = 4 
			.CommandType = 4 
			'// 存储过程名称 
			.CommandText = "GetICMaxNum" 
			'// 设置输入参数 
			set prm=.CreateParameter("TableName", 200, 1,50 , "ICJobPay")
			.Parameters.Append prm
			set prm=.CreateParameter("FInterID", 3, 2, , 1)
			.Parameters.Append prm
			set prm=.CreateParameter("Increment", 3, 1, , 1 )
			.Parameters.Append prm
			set prm=.CreateParameter("UserID", 3, 1, , k3userid)
			.Parameters.Append prm
			'// 执行存储过程 
			.Execute 
			FInterID=.Parameters.Item("FInterID").Value
			End With 
			'// 释放对象 
			Set adoComm = Nothing
			'取interid结束-----------------------------------------
			'取单号FBillNO
			FBillNO=right(year(now()),2)&"-"&right(("0"&month(now())),2)&"-"&right(("0"&day(now())),2)&"-"
			set rs2=connk3.Execute("select FBillTypeID,FFormatChar,FProjectVal,FNumMax from t_billcodeby where FBillTypeID = '700' and FFormatChar='"&FBillNO&"'")
			if rs2.eof then
				connk3.Execute("insert into t_billcodeby(FBillTypeID,FFormatChar,FProjectVal,FNumMax) values ('700','"&FBillNO&"','yy-mm-dd',2)")
				FBillNO=FBillNO&"001"
			else
				connk3.Execute("update t_billcodeby set FNumMax=FNumMax+1 where FBillTypeID = '700' and FFormatChar='"&FBillNO&"'")
				FBillNO=FBillNO&right(rs2("FNumMax")+1000,3)
			end if
			sql2="INSERT INTO ICJobPay(FInterID,FBillNo,FBrNo,FTranType,FCancellation,FStatus,FDate,FExchangeRate,FCheckDate,FBillerID,FMultiCheckDate1,FMultiCheckDate2,FMultiCheckDate3,FMultiCheckDate4,FMultiCheckDate5,FMultiCheckDate6) "
			sql2=sql2&" VALUES ("&FInterID&",'"&FBillNO&"','0',700,0,0,'"&Date()&"',1,Null,"&k3userid&",Null,Null,Null,Null,Null,Null)"
			connk3.Execute(sql2)
			FEntryID=1
			while (not rs.eof)
				'班组ID
				dim FTeamID,FWorkerID,FCOSTOBJID,FItemID,FEntrySelfR0133,FEntrySelfR0132
				dim COSTFlag:COSTFlag=false
				if rs("Bumen")="一厂" then
					FTeamID=42166
				elseif rs("Bumen")="二厂" then
					FTeamID=42165
				elseif rs("Bumen")="三厂" then
					FTeamID=42164
				elseif rs("Bumen")="眼镜布绳" then
					FTeamID=42522
				elseif rs("Bumen")="花生盒" then
					FTeamID=43015
				end if
				'操作工id
				FWorkerID=0
				set rs2=connk3.Execute("select FItemid from t_Emp where FNumber= '"&rs("EmpID")&"'")
				if not rs2.eof then
					FWorkerID=rs2("FItemid")
				end if
				'物料编号
				if rs("Product")<>"" then
					set rs2=connk3.Execute("select FItemid from t_icitem where FName= '"&rs("Product")&"'")
					if not rs2.eof then
					COSTFlag=true
					FItemID=rs2("FItemid")
					else
					FItemID=38671
					end if
				else
					FItemID=38671
				end if
				'成本对象
				FCOSTOBJID=55
				if rs("Product")<>"" and rs("OrderID")<>"" and COSTFlag then
					set rs2=connk3.Execute("select FCostObjID from icmo where fitemid="&FItemID&" and len(fmtono)>4 and left(fmtono,len(fmtono)-4)='"&rs("OrderID")&"'")
					if not rs2.eof then
						FCOSTOBJID=rs2("FCostObjID")
					end if
				end if
				FEntrySelfR0133=0
				if rs("ReasonType")<>"" then
					set rs2=connk3.Execute("select FID from t_JSLB where FName='"&rs("ReasonType")&"'")
					if not rs2.eof then
						FEntrySelfR0133=rs2("FID")
					end if
				end if
				FEntrySelfR0132=0
				set rs3=connzxpt.Execute("select top 1 Zerenbume from manusys_JishiDetails2 where SNum='"&rs("SerialNum")&"'")
				if not rs3.eof then
					set rs2=connk3.Execute("select FItemid from t_item where fitemclassid=2 and FName='"&rs3("Zerenbume")&"'")
					if not rs2.eof then
						FEntrySelfR0132=rs2("FItemid")
					end if
				end if
				sql2="INSERT INTO ICJobPayEntry (FInterID,FEntryID,FBrNo,FTeamID,FWorkerID,FDate,FICMONO,FICMOinterID,FWBNO,FFlowCardNO,FWBInterID,FOPERID,FItemID,FCOSTOBJID,FUnitID,FAuxPieceRate,FWorkauxqty,FJobPay,FTimeUnit,FSalary,FFinishTime,FHourPay,FAmount,FNote,FPercent,FProcRptInterID,FProcRptEntryID,FWorkQty,FPieceRate,FOperSN,FEntrySelfR0132,FEntrySelfR0133) "
				sql2=sql2&" VALUES ("&FInterID&","&FEntryID&",'0',"&FTeamID&","&FWorkerID&",'"&rs("RegDate")&"','','','','','',0,"&FItemID&",'"&FCOSTOBJID&"',1795,0,"&rs("Wancheng")&",0,11082,"&rs("Danjia")&","&rs("Shijian")&","&rs("Jine")&","&rs("Jine")&",'"&rs("Reason")&"',100,0,0,"&rs("Wancheng")&",0,0,"&FEntrySelfR0132&","&FEntrySelfR0133&") "
				connk3.Execute(sql2)
				FEntryID=FEntryID+1
				connzxpt.Execute("update manusys_Jishi set checkflag=4,Zhixing='"&AdminName&"',ZhixingID='"&UserName&"',ZhixingDate='"&now()&"',K3Bill='"&FBillNO&"' where SerialNum="&rs("SerialNum"))
				rs.movenext
			wend
			rs.close
			set rs=nothing
			response.Write("执行成功！")
		else
			response.Write("没有可以执行的条目，请检查！")
			response.End()
		end if
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="EmpID" then
    InfoID=request("InfoID")
		sql="select top 1 a.员工代号,a.姓名,isnull(b.JobPrice,0) as Price from [N-基本资料单头] a left join zxpt.dbo.parametersys_JobPrice b on Position=a.工作岗位 where a.员工代号 like '%"&InfoID&"%' or a.姓名 like '%"&InfoID&"%'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
				response.write ("员工不存在，或者对应工种价格不存在！")
				response.end
		else
			response.write(rs("员工代号")&"###"&rs("姓名")&"###"&rs("Price"))
		end if
		rs.close
		set rs=nothing 
  elseif detailType="SerialNum" then
			InfoID=request("InfoID")
		sql="select * from manusys_Jishi where SerialNum ="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
			if rs.bof and rs.eof then
					response.write ("对应单据不存在，请检查！")
					response.end
		else
			response.write "{""Info"":""###"",""fieldValue"":{"
			for i=0 to rs.fields.count-2
				if IsNull(rs.fields(i).value) then
			response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
			else
			response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
			end if
			next
			if IsNull(rs.fields(i).value) then
			response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&"""},")
			else
			response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&"""},")
			end if
			response.write ("""t1"":[")
			sql="select * from  manusys_JishiDetails1  where SNum ="&InfoID
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
			response.write "],""t2"":["
			sql="select * from manusys_JishiDetails2  where SNum ="&InfoID
			set rs=server.createobject("adodb.recordset")
			rs.open sql,connzxpt,1,1
			do until rs.eof
				response.write "{"
				for i=0 to rs.fields.count-1
				if IsNull(rs.fields(i).value) then
					if rs.fields(i).type="11" then
						response.write (""""&rs.fields(i).name & """:"&LCase(cstr(rs.fields(i).value))&",")
					else
						response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
					end if
				else
					if rs.fields(i).type="11" then
						response.write (""""&rs.fields(i).name & """:"&LCase(cstr(rs.fields(i).value))&",")
					else
						response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
					end if
				end if
				next
				response.write ("""bg"":""#EBF2F9""}")
				rs.movenext
			If Not rs.eof Then
				Response.Write ","
			End If
			loop
			response.write ("]}")
		end if
		rs.close
		set rs=nothing 
  elseif detailType="OrderID" then
		sql="select top "&request("limit")&" FBillNo from Seorder where FbillNo like '%"&request("q")&"%'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
'		response.write "["
		do until rs.eof
		Response.Write(rs("FBillNo"))
'		Response.Write("{""OrderID"":"""&JsonStr(rs("FBillNo"))&"""}")
	    rs.movenext
		If Not rs.eof Then
		  Response.Write vbcrlf
		End If
    loop
'		Response.Write("]")
		rs.close
		set rs=nothing 
  elseif detailType="Product" then
		sql="select top "&request("limit")&" FName from Seorder a,SEorderEntry b,t_ICitem c where a.finterid=b.finterid and b.fitemid=c.fitemid and a.FbillNo='"&request("OrderID")&"' and c.FName like '%"&request("q")&"%'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
'		response.write "["
		do until rs.eof
		Response.Write(rs("FName"))
'		Response.Write("{""OrderID"":"""&JsonStr(rs("FBillNo"))&"""}")
	    rs.movenext
		If Not rs.eof Then
		  Response.Write vbcrlf
		End If
    loop
'		Response.Write("]")
		rs.close
		set rs=nothing 
  elseif detailType="ReasonType" then
		sql="select FName from t_JSLB where FName like '%"&request("q")&"%' order by FNumber"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
'		response.write "["
		do until rs.eof
		Response.Write(rs("FName"))
'		Response.Write("{""OrderID"":"""&JsonStr(rs("FBillNo"))&"""}")
	    rs.movenext
		If Not rs.eof Then
		  Response.Write vbcrlf
		End If
    loop
'		Response.Write("]")
		rs.close
		set rs=nothing 
  elseif detailType="Users" then
    set rs = server.createobject("adodb.recordset")
    sql="select * from (select FID as pk,FNumber as val,FNumber+'/'+FName as name,0 as pspk,'' as val1,0 as Price from ICOperGroup where FDeleted=0 and FIsCurVer=1  "
    sql=sql&" union all "
    sql=sql&" select t2.FItemID pk,t3.FNumber as val,t3.FNumber+'/'+t3.FName as name,t1.FID as pspk,t3.FName as val1,isnull(b.JobPrice,0) as Price from ICOperGroup t1 inner join ICOperGroupEntry t2 on FDeleted=0 and FIsCurVer=1 and t1.FID=t2.FID inner join t_base_emp t3 on t2.FItemID =t3.FItemID left join t_submessage t4 on t4.FInterID=t3.FDuty and t4.ftypeid=29 "
		sql=sql&" left join LDERP.dbo.[N-基本资料单头] a on t3.fnumber=a.员工代号 left join zxpt.dbo.parametersys_JobPrice b on Position=a.工作岗位) aaa"
    rs.open sql,connk3,1,1
		response.Write("[")
		do until rs.eof
	%>
	{"SerialNum": "<%=rs("pk")%>", "name": "<%=rs("name")%>", "PSNum": "<%=rs("pspk")%>", "val": "<%=rs("val")%>", "val1": "<%=rs("val1")%>", "Price": "<%=rs("Price")%>"
	<%
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
elseif showType="Export" then 
	InfoID=request("InfoID")
	response.ContentType("application/vnd.ms-excel")
	response.AddHeader "Content-disposition", "attachment; filename=erpData.xls"
	sql="select a.*,b.* from manusys_Jishi a left join manusys_JishiDetails1 b on a.SerialNum=b.SNum where 1=1 "
		 if request("sd")<>"" then sql=sql&" and datediff(d,a.RegDate,'"&request("sd")&"')<=0 "
		 if request("ed")<>"" then sql=sql&" and datediff(d,a.RegDate,'"&request("ed")&"')>=0 "
		 if request("ssd")<>"" then sql=sql&" and datediff(d,a.CheckDate,'"&request("ssd")&"')<=0 "
		 if request("sed")<>"" then sql=sql&" and datediff(d,a.CheckDate,'"&request("sed")&"')>=0 "
		 if request("wt")<>"" then sql=sql&" and a.Bumen='"&request("bm")&"' "
		 if request("jg")<>"" then sql=sql&" and b.OrderID='"&request("dd")&"' "
		 if request("sh")<>"" then sql=sql&" and a.CheckFlag='"&request("sh")&"' "
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
%>
		<table width="100%" border="0" id="editDetails" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
		  <tbody id="TbDetails">
		  <tr height="24">
			<td >单号</td>
			<td >部门</td>
			<td >日期</td>
			<td >姓名</td>
			<td >订单</td>
			<td >产品型号</td>
			<td >计时类别</td>
			<td >原因</td>
			<td >开始时间</td>
			<td >结束时间</td>
			<td >所用时间</td>
			<td >完成数</td>
			<td >单价</td>
			<td >金额</td>
			<td >审核</td>
			<td >登记人</td>
			<td >登记日期</td>
			<td >段长审核</td>
			<td >厂长审核</td>
			<td >执行人</td>
		  </tr>
<%
	do until rs.eof
%>
	<tr height="24" bgcolor="#EBF2F9">
			<td ><%=rs("SerialNum")%></td>
			<td ><%=rs("Bumen")%></td>
			<td ><%=rs("RegDate")%></td>
			<td ><%=rs("EmpName")%></td>
			<td ><%=rs("OrderID")%></td>
			<td ><%=rs("Product")%></td>
			<td ><%=rs("ReasonType")%></td>
			<td ><%=rs("Reason")%></td>
			<td ><%=rs("StartTime")%></td>
			<td ><%=rs("EndTime")%></td>
			<td ><%=rs("Shijian")%></td>
			<td ><%=rs("Wancheng")%></td>
			<td ><%=rs("Danjia")%></td>
			<td ><%=rs("Jine")%></td>
			<td ><%=rs("CheckFlag")%></td>
			<td ><%=rs("Biller")%></td>
			<td ><%=rs("BillDate")%></td>
			<td ><%=rs("Checker1")%></td>
			<td ><%=rs("Checker2")%></td>
			<td ><%=rs("Checker3")%></td>
   </tr>
<%
		rs.movenext
	loop
	response.Write("</table>")
	rs.close
	set rs=nothing 
end if
 %>
