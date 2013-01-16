<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
if Instr(session("AdminPurview"),"|1104,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限'
dim showType,start_date,end_date,print_tag,UserName,AdminName
UserName=session("UserName")
AdminName=session("AdminName")
showType=request("showType")
print_tag=request("print_tag")
if showType="DetailsList" then 
elseif showType="DataProcess" then 
  detailType=request("detailType")
  if detailType="AddNew" then
	set rs = server.createobject("adodb.recordset")
	sql="select * from smmsys_Permission"
	rs.open sql,connzxpt,1,3
	rs.addnew
	rs("PermissionClass")=Request("PermissionClass")
	rs("LongName")=Request("LongName")
	rs("PSNum")=Request("PSNum")
	rs("PermissionName")="新增"
	if Request("N_lv")>=1 then
	  set rs2 = server.createobject("adodb.recordset")
	  sql2="select count(1) as idcount from smmsys_Permission where PSNum="&Request("PSNum")
	  rs2.open sql2,connzxpt,1,1
	  newidsum=rs2("idcount")+1
	  rs2.close
	  set rs2=nothing
	  rs("PermissionID")=left(Request("PermissionID"),len(Request("PermissionID"))-1)&"."&newidsum&","
	else
	  rs("PermissionID")="|,"
	end if
	rs.update
	set rs = server.createobject("adodb.recordset")
	sql="select top 1 * from smmsys_Permission order by serialnum desc"
	rs.open sql,connzxpt,1,1
	response.write("[{""SerialNum"": """&rs("SerialNum")&""", ""PermissionID"": """&rs("PermissionID")&""", ""name"": """&rs("PermissionName")&""", ""PSNum"": """&rs("PSNum")&""", ""LongName"": """&rs("LongName")&"""}]")
	rs.close
	set rs=nothing 
  elseif detailType="Edit" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from smmsys_Permission where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据连接失败，请检查！")
		response.end
	end if
	rs("PermissionID")=Request("PermissionID")
	rs("PermissionName")=Request("PermissionName")
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif detailType="SetEdit" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("EmpId")
    sql="select * from [N-基本资料单头] where 员工代号='"&SerialNum&"'"
    rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		response.write ("员工编号不存在，请检查！")
		response.end
	end if
	if request("PurviewType")="FLW" then
	rs("权限_FLW")=request("PermissionID")
	else
	rs("权限")=request("PermissionID")
	end if
	rs.update
	rs.close
	set rs=nothing 
	response.write "保存成功，对应员工重新登录才能生效！"
  elseif detailType="Delete" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from smmsys_Permission where SerialNum="&SerialNum
	rs.open sql,connzxpt,0,1
	if rs.bof and rs.eof then
		response.write ("数据连接失败，请检查！")
		response.end
	end if
	set rs = server.createobject("adodb.recordset")
	sql="select count(1) as idcount from smmsys_Permission where PSNum="&SerialNum
	rs.open sql,connzxpt,0,1
	if rs("idcount")>0 then
		response.write ("该权限下面有子节点，不允许删除！")
		response.end
	end if
	rs.close
	set rs=nothing 
	connzxpt.Execute("Delete from smmsys_Permission where SerialNum="&SerialNum)
	response.write "###"
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="Register" then
    InfoID=request("InfoID")
	sql="select a.姓名,a.部门别,b.部门名称 from [N-基本资料单头] a,[G-部门资料表] b where a.部门别=b.部门代号 and a.员工代号='"&InfoID&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,0,1
    if rs.bof and rs.eof then
        response.write ("员工编号不存在！")
        response.end
	else
		response.write("###"&rs("姓名")&"###"&rs("部门别")&"###"&rs("部门名称")&"###")
	end if
	rs.close
	set rs=nothing 
  elseif detailType="SetPurview" then
	dim tempstr1,EmpId
	tempstr1=""
	PermissionID=""
	EmpId=request("EmpId")

	if request("PurviewType")="RPT" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from [N-基本资料单头] where 员工代号='"&EmpId&"'"
      rs.open sql,conn,1,1
	  PermissionID=rs("权限")
		response.Write("[")
		sql="select * from smmsys_Permission where PermissionClass='RPT' order by longname,permissionid asc"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		do until rs.eof
	%>
	{"SerialNum": "<%=rs("SerialNum")%>", "PermissionID": "<%=rs("PermissionID")%>", "name": "<%=rs("PermissionName")%>", "PSNum": "<%=rs("PSNum")%>"
	<%
	if instr(PermissionID,rs("PermissionID"))>0 then response.Write(",checked : true")
	
		set rs2=server.createobject("adodb.recordset")
		sql2="select count(1) as idcount from smmsys_Permission where PermissionClass='RPT' and PSNum="&rs("SerialNum")
		rs2.open sql2,connzxpt,1,1
		if rs2("idcount")>0 then
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
		rs2.close
		set rs2=nothing 
	elseif request("PurviewType")="FLW" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from [N-基本资料单头] where 员工代号='"&EmpId&"'"
      rs.open sql,conn,1,1
	  PermissionID=rs("权限_FLW")
		response.Write("[")
		if InfoID="" then
		  tempstr1=" and psnum=0"
		else
		  tempstr1=" and psnum="&InfoID
		end if
		sql="select * from smmsys_Permission where PermissionClass='FLW' order by longname,permissionid asc"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		do until rs.eof
	%>
	{"SerialNum": "<%=rs("SerialNum")%>", "PermissionID": "<%=rs("PermissionID")%>", "name": "<%=rs("PermissionName")%>", "PSNum": "<%=rs("PSNum")%>"
	<%
	if instr(PermissionID,rs("PermissionID"))>0 then response.Write(",checked : true")
	
		set rs2=server.createobject("adodb.recordset")
		sql2="select count(1) as idcount from smmsys_Permission where PermissionClass='FLW' and PSNum="&rs("SerialNum")
		rs2.open sql2,connzxpt,1,1
		if rs2("idcount")>0 then
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
		rs2.close
		set rs2=nothing 
	end if
	rs.close
	set rs=nothing 
  elseif detailType="EditPurview" then
    InfoID=request("SerialNum")
	response.Write("[")
	if InfoID="" then
	  tempstr1=" and psnum=0"
	else
	  tempstr1=" and psnum="&InfoID
	end if
	sql="select * from smmsys_Permission where PermissionClass='"&request("PurviewType")&"' "&tempstr1&" order by longname,permissionid asc"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
	do until rs.eof
	%>
	{"SerialNum": "<%=rs("SerialNum")%>", "PermissionID": "<%=rs("PermissionID")%>", "name": "<%=rs("PermissionName")%>", "PSNum": "<%=rs("PSNum")%>", "LongName": "<%=rs("LongName")%>"
	<%
		set rs2=server.createobject("adodb.recordset")
		sql2="select count(1) as idcount from smmsys_Permission where PermissionClass='"&request("PurviewType")&"' and PSNum="&rs("SerialNum")
		rs2.open sql2,connzxpt,1,1
		if rs2("idcount")>0 then
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
		rs2.close
		set rs2=nothing 
	rs.close
	set rs=nothing 
  
  end if
end if
 %>
