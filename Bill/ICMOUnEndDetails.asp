<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurviewFLW"),"|212,")=0 then 
'  response.write "{""page"":""1"",""total"":""1"",""rows"":[{""id"":""0"",""cell"":["
'  response.write ("""<font color='red')>你不具有该管理模块的操作权限，请返回！</font>""]}]}")
'  response.end
'end if
'========判断是否具有管理权限
dim showType,start_date,end_date,print_tag,UserName,AdminName,Depart
Depart=session("Depart")
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
      datafrom=" Bill_ICMOUnEnd "
  dim datawhere'数据条件
  dim i'用于循环的整数
    datawhere=" where 1=1 "&Session("AllMessage21")&Session("AllMessage22")&Session("AllMessage23")&Session("AllMessage24")
		Session("AllMessage21")=""
		Session("AllMessage22")=""
		Session("AllMessage23")=""
		Session("AllMessage24")=""
	if request("AllQuery")<>"" then
		datawhere=datawhere&" and "&request("AllQuery")
	else
		datawhere=datawhere&" and CancelFlag=0 "
	end if
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
    do until rs.eof'填充数据到表格
		dim tempstr2:tempstr2="未审核"
		if rs("CheckFlag")=1 then
		tempstr2="已审核"
		elseif rs("CheckFlag")=2 then
		tempstr2="已审批"
		elseif rs("CheckFlag")=3 then
		tempstr2="已执行"
		elseif rs("CheckFlag")=4 then
		tempstr2="已结案"
		end if

%>
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("RegDate")%>","<%=rs("Departmentname")%>","<%=rs("RegisterName")%>","<%=rs("ICMOId")%>","<%=rs("OrderId")%>","<%=rs("RegistType")%>","<%=replace(replace(rs("Reason"),chr(10),"\n"),chr(13),"\r")%>","<%=rs("PlanEndDate")%>","<%=rs("EndDate")%>","<%=tempstr2%>","<%=rs("CancelFlag")%>","<%=replace(replace(rs("Remark"),chr(10),"\n"),chr(13),"\r")%>"]}
<%	
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
  end if
  rs.close
  set rs=nothing
response.Write "]}"
'-----------------------------------------------------------
elseif showType="DataProcess" then 
  detailType=request("detailType")
  if detailType="AddNew" then
'  	if  Instr(session("AdminPurviewFLW"),"|212.1,")>0 then
		set rs = server.createobject("adodb.recordset")
		sql="select * from Bill_ICMOUnEnd"
		rs.open sql,connzxpt,1,3
		rs.addnew
		rs("Register")=request("Register")
		rs("RegisterName")=request("RegisterName")
		rs("RegDate")=request("RegDate")
		rs("Department")=request("Department")
		rs("Departmentname")=request("Departmentname")
		rs("ICMOId")=request("ICMOId")
		rs("OrderId")=request("OrderId")
		rs("RegistType")=request("RegistType")
		rs("Reason")=request("Reason")
		rs("Remark")=request("Remark")
		rs("Biller")=AdminName
		rs("BillDate")=now()
		if request("PlanEndDate")<>"" then rs("PlanEndDate")=request("PlanEndDate")
		rs.update
		rs.close
		set rs=nothing 
		response.write "###"
'	else
'		response.write ("你没有权限进行此操作！")
'		response.end
'	end if
  elseif detailType="Edit"  then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
		sql="select * from Bill_ICMOUnEnd where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if rs("Biller")<>AdminName and rs("Register")<>UserName then
			response.write ("只能编辑自己的单据！")
			response.end
    end if
		if rs("CheckFlag")>0 then
			response.write ("已审核不允许编辑！")
			response.end
    end if
		rs("Register")=request("Register")
		rs("RegisterName")=request("RegisterName")
		rs("RegDate")=request("RegDate")
		rs("Department")=request("Department")
		rs("Departmentname")=request("Departmentname")
		rs("ICMOId")=request("ICMOId")
		rs("OrderId")=request("OrderId")
		rs("RegistType")=request("RegistType")
		rs("Reason")=request("Reason")
		rs("Remark")=request("Remark")
		rs("Biller")=AdminName
		rs("BillDate")=now()
		if request("PlanEndDate")<>"" then rs("PlanEndDate")=request("PlanEndDate")
    rs.update
		rs.close
		set rs=nothing 
		response.write "###"
  elseif detailType="Delete" then
		set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
		sql="select * from Bill_ICMOUnEnd where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,1
		while (not rs.eof)
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if rs("Biller")<>AdminName and rs("Register")<>UserName then
			response.write ("只能删除自己的单据！")
			response.end
    end if
		if rs("CheckFlag")>0 then
			response.write ("单据已审核，不允许删除！")
			response.end
		end if
			rs.movenext
		wend
		rs.close
		set rs=nothing 
		connzxpt.Execute("Delete from Bill_ICMOUnEnd where SerialNum in ("&SerialNum&")")
		response.write "###"
  elseif detailType="审核" then
		dim direct:direct=request("direct")
    SerialNum=request("SerialNum")
		set rs = server.createobject("adodb.recordset")
		sql="select * from Bill_ICMOUnEnd where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurviewFLW"),"|212.2,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		if rs("CheckFlag")>1 then
			response.write ("单据当前状态不允许此操作！")
			response.end
		end if
		if rs("Department")<>Depart then
		
			response.write ("只能审核本部门的单据！")
			response.end
		end if
		rs("Checker")=AdminName
		rs("CheckDate")=now()
		rs("CheckText")=request("operattext")
		if direct="Y" then
		rs("CheckFlag")=1
		elseif direct="N" then
		rs("CheckFlag")=0
		rs("CancelFlag")=0
		elseif direct="Z" then
		rs("CheckFlag")=1
		rs("CancelFlag")=1
		end if
			rs.movenext
		wend
		rs.update
		rs.close
		set rs=nothing 
		response.write "审核成功！"
  elseif detailType="审批" then
		direct=request("direct")
    SerialNum=request("SerialNum")
		set rs = server.createobject("adodb.recordset")
		sql="select * from Bill_ICMOUnEnd where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurviewFLW"),"|212.3,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		if rs("CheckFlag")>2 or rs("CheckFlag")<1 or rs("CancelFlag") then
			response.write ("单据当前状态不允许此操作！")
			response.end
		end if
		rs("Approvaler")=AdminName
		rs("ApprovalDate")=now()
		rs("ApprovalText")=request("operattext")
		if direct="Y" then
		rs("CheckFlag")=2
		elseif direct="N" then
		rs("CheckFlag")=1
		rs("CancelFlag")=0
		elseif direct="Z" then
		rs("CheckFlag")=2
		rs("CancelFlag")=1
		end if
			rs.movenext
		wend
		rs.update
		rs.close
		set rs=nothing 
		response.write "审批成功！"
  elseif detailType="执行" then
		direct=request("direct")
    SerialNum=request("SerialNum")
		set rs = server.createobject("adodb.recordset")
		sql="select * from Bill_ICMOUnEnd where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurviewFLW"),"|212.4,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		if rs("CheckFlag")>3 or rs("CheckFlag")<2 or rs("CancelFlag") then
			response.write ("单据当前状态不允许此操作！")
			response.end
		end if
		rs("Implementer")=AdminName
		rs("ImplementDate")=now()
		rs("ImplementText")=request("operattext")
		if direct="Y" then
		rs("CheckFlag")=3
		elseif direct="N" then
		rs("CheckFlag")=2
		rs("CancelFlag")=0
		elseif direct="Z" then
		rs("CheckFlag")=3
		rs("CancelFlag")=1
		end if
			rs.movenext
		wend
		rs.update
		rs.close
		set rs=nothing 
		response.write "执行成功！"
  elseif detailType="结案" then
		direct=request("direct")
    SerialNum=request("SerialNum")
		set rs = server.createobject("adodb.recordset")
		sql="select * from Bill_ICMOUnEnd where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurviewFLW"),"|212.4,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		if rs("CheckFlag")<3 or rs("CancelFlag") then
			response.write ("单据当前状态不允许此操作！")
			response.end
		end if
		rs("Implementer")=AdminName
		rs("ImplementDate")=now()
		rs("ImplementText")=request("operattext")
		if direct="Y" then
		rs("CheckFlag")=4
		rs("EndDate")=date()
		elseif direct="N" then
		rs("CheckFlag")=3
		rs("CancelFlag")=0
		rs("EndDate")=null
		elseif direct="Z" then
		rs("CheckFlag")=4
		rs("CancelFlag")=1
		end if
			rs.movenext
		wend
		rs.update
		rs.close
		set rs=nothing 
		response.write "结案成功！"
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="Register" or detailType="Zeren" then
    InfoID=request("InfoID")
		if InfoID="" then InfoID=UserName
		sql="select a.*,b.部门名称 from [N-基本资料单头] a,[G-部门资料表] b where a.部门别=b.部门代号 and (a.员工代号 like '%"&InfoID&"%' or a.姓名 like '%"&InfoID&"%')"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
			if rs.bof and rs.eof then
					response.write ("员工编号不存在！")
					response.end
		else
			response.write "{""Info"":""###"",""fieldValue"":[{""Register"":"""&rs("员工代号")&""",""RegisterName"":"""&rs("姓名")&""",""Departmentname"":"""&rs("部门名称")&""",""Department"":"""&rs("部门别")&"""}]}"
		end if
		rs.close
		set rs=nothing 
  elseif detailType="ICMOId" then
    InfoID=request("InfoID")
		sql="select a.*,b.FbillNo as orderid,c.FName as FDepartMent,d.FNumber,d.Fname,d.FModel  from ICMO a left join Seorder b on a.forderinterid=b.FInterID inner join t_department c on a.FWorkShop=c.Fitemid inner join t_icitem d on a.Fitemid=d.Fitemid where a.FBillNo='"&InfoID&"'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
			if rs.bof and rs.eof then
					response.write ("生产单编号不存在！")
					response.end
		else
			dim temstr1:temstr1="未结案"
			if rs("FStatus")=3 then temstr1="已结案"
			response.write "{""Info"":""###"",""fieldValue"":[{""ICMOId"":"""&rs("FBillno")&""",""OrderId"":"""&rs("orderid")&""",""FNumber"":"""&rs("FNumber")&""",""FName"":"""&rs("FName")&""",""FModel"":"""&rs("FModel")&""",""FDepartMent"":"""&rs("FDepartMent")&""",""FAuxQty"":"""&rs("FAuxQty")&""",""FPlanFinishDate"":"""&rs("FPlanFinishDate")&""",""FAuxStockQty"":"""&rs("FAuxStockQty")&""",""FStatus"":"""&temstr1&"""}]}"
		end if
		rs.close
		set rs=nothing 
  elseif detailType="SerialNum" then
    InfoID=request("InfoID")
		sql="select e.*,b.Fname as FDepartMent,d.FNumber,d.FName,d.FModel,a.FAuxQty,a.FPlanFinishDate,a.FAuxStockQty,Case when a.FStatus=4 then '已结案' else '未结案' end as FStatus from Bill_ICMOUnEnd e left join "&AllOPENROWSET&" AIS20081217153921.dbo.ICMO) as a on a.FBillno=e.ICMOId left join  "&AllOPENROWSET&" AIS20081217153921.dbo.t_department) as b on a.FWorkShop=b.Fitemid left join  "&AllOPENROWSET&" AIS20081217153921.dbo.t_icitem) as d on a.Fitemid=d.Fitemid where SerialNum="&InfoID
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
  elseif detailType="审核" then
    InfoID=request("InfoID")
		sql="select * from Bill_ICMOUnEnd where SerialNum="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
			if rs.bof and rs.eof then
					response.write ("单据编号不存在！")
					response.end
		else
			response.write "{""Info"":""###"",""operater"":"""&rs("Checker")&""",""operattime"":"""&rs("CheckDate")&""",""operattext"":"""&replace(replace(rs("CheckText"),chr(10),"\n"),chr(13),"\r")&"""}"
		end if
		rs.close
		set rs=nothing 
  elseif detailType="审批" then
    InfoID=request("InfoID")
		sql="select * from Bill_ICMOUnEnd where SerialNum="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
			if rs.bof and rs.eof then
					response.write ("单据编号不存在！")
					response.end
		else
			response.write "{""Info"":""###"",""operater"":"""&rs("Approvaler")&""",""operattime"":"""&rs("ApprovalDate")&""",""operattext"":"""&replace(replace(rs("ApprovalText"),chr(10),"\n"),chr(13),"\r")&"""}"
		end if
		rs.close
		set rs=nothing 
  elseif detailType="执行" or detailType="结案" then
    InfoID=request("InfoID")
		sql="select * from Bill_ICMOUnEnd where SerialNum="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
			if rs.bof and rs.eof then
					response.write ("单据编号不存在！")
					response.end
		else
			response.write "{""Info"":""###"",""operater"":"""&rs("Implementer")&""",""operattime"":"""&rs("ImplementDate")&""",""operattext"":"""&replace(replace(rs("ImplementText"),chr(10),"\n"),chr(13),"\r")&"""}"
		end if
		rs.close
		set rs=nothing 
  elseif detailType="Excel" then
		if request("AllQuery")<>"" then
			datawhere=" and "&request("AllQuery")
		end if
		sql="select * from Bill_ICMOUnEnd where 1=1 "&datawhere
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
%>
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>ID</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请部门</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>生产任务单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>订单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>原因分类</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>反结案原因</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>预计完成时间</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>实际完成时间</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>状态</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>备注</strong></font></td>
  </tr>
<%
		tempstr2="未审核"
		if rs("CheckFlag")=1 then
		tempstr2="已审核"
		elseif rs("CheckFlag")=2 then
		tempstr2="已审批"
		elseif rs("CheckFlag")=3 then
		tempstr2="已执行"
		elseif rs("CheckFlag")=4 then
		tempstr2="已结案"
		end if
		while(not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("SerialNum")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("RegDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("RegisterName")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ICMOId")&"</td>"
      Response.Write "<td nowrap>"&rs("OrderId")&"</td>"
      Response.Write "<td nowrap>"&rs("RegistType")&"</td>"
      Response.Write "<td nowrap>"&rs("Reason")&"</td>"
      Response.Write "<td nowrap>"&rs("PlanEndDate")&"</td>"
      Response.Write "<td nowrap>"&rs("EndDate")&"</td>"
      Response.Write "<td nowrap>"&tempstr2&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Remark")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
			rs.movenext
		wend
		Response.Write "</table>" & vbCrLf
		rs.close
		set rs=nothing 
  end if
end if
 %>
