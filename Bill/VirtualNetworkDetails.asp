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
      datafrom=" Bill_VirtualNetwork "
  dim datawhere'数据条件
  dim i'用于循环的整数
	if Instr(session("AdminPurviewFLW"),"|219.1,")>0 then
    datawhere=" where 1=1 "
'	elseif Instr(session("AdminPurviewFLW"),"|219.1,")>0 and Depart="KD01.0001.0010" then
'		datawhere=" where (Department='"&Depart&"' or Department='KD01.0001.0018' or Biller='"&AdminName&"') "
	else
		datawhere=" where (Department='"&Depart&"' or BillerID='"&UserName&"' or RegisterID='"&UserName&"') "
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
	datawhere=datawhere&Session("AllMessage59")&Session("AllMessage71")
	session.contents.remove "AllMessage59"
	session.contents.remove "AllMessage71"
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
			dim tempstr,tempstr2
			if rs("CheckFlag")=0 then
				tempstr="未执行"
			elseif rs("CheckFlag")=1 then
				tempstr="已执行"
			elseif rs("CheckFlag")=2 then
				tempstr="已结案"
			end if
			
			if Instr(session("AdminPurviewFLW"),"|219.1,")>0 or rs("RegisterID")=UserName or rs("BillerID")=UserName then
				tempstr2=rs("ServicePwd")
			else
				tempstr2="******"
			end if
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("RegDate")%>","<%=rs("RegisterID")%>","<%=rs("Register")%>","<%=rs("Departmentname")%>","<%=rs("Telephone")%>","<%=rs("ModifyType")%>","<%=rs("PackageNote")%>","<%=rs("Enforcement")%>","<%=tempstr2%>","<%=tempstr%>","<%=rs("CancelFlag")%>","<%=rs("NewNumber")%>","<%=rs("Biller")%>","<%=rs("Checker")%>"]}
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
		sql="select * from Bill_VirtualNetwork"
		rs.open sql,connzxpt,1,3
		rs.addnew
		rs("BillerID")=UserName
		rs("Biller")=AdminName
		rs("BillDate")=now()
		rs("Register")=Request("Register")
		rs("RegisterID")=Request("RegisterID")
		rs("RegDate")=Request("RegDate")
		rs("Department")=Request("Department")
		rs("Departmentname")=Request("Departmentname")
		rs("Telephone")=Request("Telephone")
		rs("ModifyType")=Request("ModifyType")
		rs("PackageNote")=Request("PackageNote")
		rs("Enforcement")=Request("Enforcement")
		rs("ServicePwd")=Request("ServicePwd")
		rs.update
		rs.close
		set rs=nothing 
		response.write "###"
  elseif detailType="Edit"  then
		set rs = server.createobject("adodb.recordset")
		sql="select * from Bill_VirtualNetwork where SerialNum="&request("SerialNum")
		rs.open sql,connzxpt,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if rs("CheckFlag")>0 then
			response.write ("当前状态不允许编辑，请检查！")
			response.end
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
		rs("RegDate")=Request("RegDate")
		rs("Department")=Request("Department")
		rs("Departmentname")=Request("Departmentname")
		rs("Telephone")=Request("Telephone")
		rs("ModifyType")=Request("ModifyType")
		rs("PackageNote")=Request("PackageNote")
		rs("Enforcement")=Request("Enforcement")
		rs("ServicePwd")=Request("ServicePwd")
		rs.update
		rs.close
		set rs=nothing 
		response.write "###"
	elseif detailType="Check" then
		set rs = server.createobject("adodb.recordset")
			SerialNum=request("SerialNum")
		sql="select * from Bill_VirtualNetwork where SerialNum ="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs("CheckFlag")=0 then
			if Instr(session("AdminPurviewFLW"),"|219.1,")=0 then
				response.write ("你没有权限进行当前操作！")
				response.end
			end if
			rs("CheckerID")=UserName
			rs("Checker")=AdminName
			rs("CheckDate")=now()
			rs("CheckFlag")=1
			rs("CancelFlag")=Request("flag")
		end if
		rs.update
		rs.close
		set rs=nothing 
		response.write("###")
	elseif detailType="Finish" then
		set rs = server.createobject("adodb.recordset")
			SerialNum=request("SerialNum")
		sql="select * from Bill_VirtualNetwork where SerialNum ="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs("BillerID")<>UserName and rs("RegisterID")<>UserName then
			response.write ("你没有权限进行当前操作！")
			response.end
		end if
		if rs("CheckFlag")>0 then
			if Request("flag")="0" then
				rs("CheckFlag")=1
			else
				if rs("ModifyType")="新增加入" or rs("ModifyType")="修改短号" then
					if Request("NewNumber")="" then
						response.Write("请先填写新短号，再进行结案！")
						response.end
					end if
					rs("NewNumber")=Request("NewNumber")
					connk3.Execute("update t_Base_Emp set shortMobile='"&Request("NewNumber")&"' where FNumber ='"&rs("RegisterID")&"'")
				end if
				rs("CheckFlag")=2
			end if
		end if
		rs.update
		rs.close
		set rs=nothing 
		response.write("###")
  elseif detailType="Delete" then
    SerialNum=request("SerialNum")
		sql="select * from Bill_VirtualNetwork where CheckFlag>0 and SerialNum in ("&SerialNum&")"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		if rs.eof then
			connzxpt.Execute("Delete from Bill_VirtualNetwork where SerialNum in ("&SerialNum&")")
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
			response.write(rs("员工代号")&"###"&rs("姓名")&"###"&rs("部门别")&"###"&rs("部门名称")&"###"&rs("手机全号"))
		end if
		rs.close
		set rs=nothing 
  elseif detailType="SerialNum" then
    InfoID=request("InfoID")
		sql="select * from Bill_VirtualNetwork where SerialNum="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		if rs.bof and rs.eof then
			response.write ("对应单据不存在，请检查！")
			response.end
		else
			response.write "{""Info"":""###"",""fieldValue"":[{"
			if Instr(session("AdminPurviewFLW"),"|219.1,")>0 or rs("RegisterID")=UserName or rs("BillerID")=UserName then
				tempstr2=rs("ServicePwd")
			else
				tempstr2="******"
			end if
%>
"SerialNum":"<%=rs("SerialNum")%>",
"RegDate":"<%=rs("RegDate")%>",
"RegisterID":"<%=rs("RegisterID")%>",
"Register":"<%=rs("Register")%>",
"Department":"<%=rs("Department")%>",
"Departmentname":"<%=rs("Departmentname")%>",
"Telephone":"<%=rs("Telephone")%>",
"ModifyType":"<%=rs("ModifyType")%>",
"PackageNote":"<%=rs("PackageNote")%>",
"Enforcement":"<%=rs("Enforcement")%>",
"ServicePwd":"<%=tempstr2%>"}]}
<%
		end if
		rs.close
		set rs=nothing 
	end if
end if
 %>
