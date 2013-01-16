<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurviewFLW"),"|203,")=0 then 
'  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
'  response.end
'end if
'========判断是否具有管理权限
dim showType,start_date,end_date,print_tag,UserName,AdminName,GroupName,GroupID,detailType
UserName=session("UserName")
AdminName=session("AdminName")
showType=request("showType")
print_tag=request("print_tag")
dim sql,rs
if print_tag=1 then
response.ContentType("application/vnd.ms-excel")
response.AddHeader "Content-disposition", "attachment; filename=erpData.xls"
end if
if showType="DetailsList" then 

elseif showType="AddEditShow" then 
  detailType=request("detailType")
  dim Purview:Purview=""
  GroupID=request("GroupID")
  set rs = server.createobject("adodb.recordset")
  sql="select * from smmsys_PermissionGroup where SerialNum="&GroupID
  rs.open sql,connzxpt,1,1
  GroupName=rs("GroupName")
  rs.close
  set rs=nothing 

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HEAD>
<TITLE>欢迎进入系统后台</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2011-2012" />
<META NAME="Author" CONTENT="报表系统" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<link rel="stylesheet" href="../Images/CssAdmin.css">

</HEAD>
<BODY>
 <div id="AddandEditdiv" style="width:790px; height:460px; overflow:auto; ">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <form id="editForm" name="editForm" method="post" >
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead1()" >&nbsp;<strong>绩效报表系统组别权限设置</strong></font></td>
  </tr>
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">组&nbsp;别&nbsp;名：</td>
        <td>
		<input type="hidden" name="GroupID" id="GroupID" value="<%=GroupID%>">
		<input name="UserName" type="text" class="textfield" id="UserName" style="WIDTH: 120;" value="<%=GroupName%>" maxlength="16" readonly>&nbsp;*&nbsp;3-10位字符，不可修改
		</td>
      </tr>
      <tr >
        <td height="20" align="right">操作权限：</td>
        <td nowrap>
  		<div class="zTreeDemoBackground">
			<ul id="PermissionRPT" class="tree"></ul>
		</div>		
		  </td>
      <tr>
        <td height="30" align="right">&nbsp;</td>
        <td valign="bottom">
		<input name="submitSaveEdit" type="button" class="button"  id="submitSaveEdit" value="保存" onClick="submitSave()" style="WIDTH: 60;" >
		<input type="hidden" name="detailType" id="detailType" value="<%=detailType%>">
		</td>
      </tr>
      <tr>
        <td height="20" align="right">&nbsp;</td>
        <td valign="bottom">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  </form>
</table>

 </div>
</body>
</html>
<%
elseif showType="DataProcess" then 
  detailType=request("detailType")
  if detailType="AddNew" then
    GroupName=request("GroupName")
	set rs = server.createobject("adodb.recordset")
	sql="select * from smmsys_PermissionGroup where GroupName='"&GroupName&"'"
	rs.open sql,connzxpt,0,1
	if rs.eof and rs.bof then
	  connzxpt.Execute("insert into smmsys_PermissionGroup (GroupName) values ('"&GroupName&"')")
	  set rs = server.createobject("adodb.recordset")
	  sql="select * from smmsys_PermissionGroup where GroupName='"&GroupName&"'"
	  rs.open sql,connzxpt,0,1
	  response.write "###"&rs("SerialNum")&"###"&rs("GroupName")&"###"
	else
	  response.write "@@@该组名已存在，不能重复添加！@@@"
	  response.End()
	end if
	rs.close
	set rs=nothing 
  elseif detailType="Edit" then
	set rs = server.createobject("adodb.recordset")
    GroupName=request("GroupName")
	GroupID=request("GroupID")
	sql="select * from smmsys_PermissionGroup where GroupName='"&GroupName&"' and SerialNum!="&GroupID
	rs.open sql,connzxpt,0,1
	if rs.eof and rs.bof then
	  connzxpt.Execute("update smmsys_PermissionGroup set GroupName='"&GroupName&"' where SerialNum="&GroupID)
	  set rs = server.createobject("adodb.recordset")
	  sql="select * from smmsys_PermissionGroup where GroupName='"&GroupName&"'"
	  rs.open sql,connzxpt,0,1
	  response.write "###"&GroupID&"###"&GroupName&"###"
	else
	  response.write "@@@该组名已存在，不能重复添加！@@@"
	  response.End()
	end if
	rs.close
	set rs=nothing 
  elseif detailType="Delete" then
	set rs = server.createobject("adodb.recordset")
    GroupID=request("GroupID")
	sql="select * from smmsys_PermissionGroup where SerialNum="&GroupID
	rs.open sql,connzxpt,0,1
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	rs.close
	set rs=nothing 
	connzxpt.Execute("Delete from smmsys_PermissionGroup where SerialNum="&GroupID)
	connzxpt.Execute("Delete from smmsys_PermissionGroupDetails where GroupSnum="&GroupID)
	response.write "###"&GroupID&"###"
  elseif detailType="Forbid" then
	set rs = server.createobject("adodb.recordset")
    GroupID=request("GroupID")
	sql="select * from smmsys_PermissionGroup where SerialNum="&GroupID
	rs.open sql,connzxpt,0,1
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	rs.close
	set rs=nothing 
	connzxpt.Execute("update smmsys_PermissionGroup set ForbidFlag=1 where SerialNum="&GroupID)
	response.write "###"&GroupID&"###"
  elseif detailType="unForbid" then
	set rs = server.createobject("adodb.recordset")
    GroupID=request("GroupID")
	sql="select * from smmsys_PermissionGroup where SerialNum="&GroupID
	rs.open sql,connzxpt,0,1
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	rs.close
	set rs=nothing 
	connzxpt.Execute("update smmsys_PermissionGroup set ForbidFlag=0 where SerialNum="&GroupID)
	response.write "###"&GroupID&"###"
  elseif detailType="AddUser" then
    GroupID=request("GroupID")
	dim UserIDs:UserIDs=request("UserIDs")
	dim UserID:UserID=Split(UserIDs, ",")
	set rs = server.createobject("adodb.recordset")
	sql="select * from smmsys_PermissionGroup where SerialNum="&GroupID
	rs.open sql,connzxpt,0,1
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	dim i:i=0
	while(i<=UBound(UserID))
	  connzxpt.Execute("insert into smmsys_PermissionGroupDetails (GroupSnum,UserID) values("&GroupID&",'"&UserID(i)&"')")
	  i=i+1
	wend
	rs.close
	set rs=nothing 
	response.write "###"&GroupID&"###"
  elseif detailType="DeleteUser" then
    GroupID=request("GroupID")
	UserIDs=request("UserIDs")
	UserID=Split(UserIDs, ",")
	set rs = server.createobject("adodb.recordset")
	sql="select * from smmsys_PermissionGroup where SerialNum="&GroupID
	rs.open sql,connzxpt,0,1
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	i=0
	while(i<=UBound(UserID))
	  connzxpt.Execute("delete from smmsys_PermissionGroupDetails where GroupSnum="&GroupID&" and UserID='"&UserID(i)&"'")
	  i=i+1
	wend
	rs.close
	set rs=nothing 
	response.write "###"&GroupID&"###"
		sql="select a.员工代号 as pk,a.员工代号+'/'+a.姓名 as name,a.部门别 as pspk from [N-基本资料单头] a  where a.离职否='在职' and 职等!='001' and 是否离职=0 "
		sql=sql&" union all  "
		sql=sql&" select distinct c.部门代号 as pk,c.部门名称 as name,'0' as pspk from [N-基本资料单头] a,[G-部门资料表] c  where a.离职否='在职' and a.部门别=c.部门代号 "
		sql=sql&" except "
		sql=sql&" select a.UserID as pk,b.员工代号+'/'+b.姓名 as name,b.部门别 as pspk from  zxpt.dbo.smmsys_PermissionGroupDetails a,[N-基本资料单头] b where b.员工代号=a.UserID and a.GroupSnum="&GroupID
		set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
		response.Write("[")
		do until rs.eof
	%>
	{"SerialNum": "<%=rs("pk")%>", "name": "<%=rs("name")%>", "PSNum": "<%=rs("pspk")%>","val1":"<%=rs("val1")%>","val2":"<%=rs("val2")%>","val3":"<%=rs("val3")%>","val4":"<%=rs("val4")%>"
	<%
		if rs("pspk")="0" then
	%>
	, "isParent":"true"
	<%
		end if
		response.Write("}")
			rs.movenext
			If Not rs.eof Then
			  Response.Write ","
			End If
		loop
		Response.Write "]"
  elseif detailType="AddAllUser" then
    GroupID=request("GroupID")
	set rs = server.createobject("adodb.recordset")
	sql="select * from smmsys_PermissionGroup where SerialNum="&GroupID
	rs.open sql,connzxpt,0,1
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	set rs = server.createobject("adodb.recordset")
	sql="select 员工代号 from [N-基本资料单头] where 职等!='001' and 是否离职=0 EXCEPT select a.UserID from  "&AllOPENROWSET&" zxpt.dbo.smmsys_PermissionGroupDetails) a where a.GroupSnum="&GroupID
	rs.open sql,conn,0,1
	while(not rs.eof)
	  connzxpt.Execute("insert into smmsys_PermissionGroupDetails (GroupSnum,UserID) values("&GroupID&",'"&rs("员工代号")&"')")
	  rs.movenext
	wend
	rs.close
	set rs=nothing 
	response.write "###"&GroupID&"###"
  elseif detailType="DeleteAllUser" then
    GroupID=request("GroupID")
	set rs = server.createobject("adodb.recordset")
	sql="select * from smmsys_PermissionGroup where SerialNum="&GroupID
	rs.open sql,connzxpt,0,1
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	connzxpt.Execute("delete from smmsys_PermissionGroupDetails where GroupSnum="&GroupID)
	rs.close
	set rs=nothing 
	response.write "###"&GroupID&"###"
  elseif detailType="SetEdit" then
    GroupID=request("GroupID")
	
	set rs = server.createobject("adodb.recordset")
	sql="select * from smmsys_PermissionGroup where SerialNum="&GroupID
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if Request("PurviewType")="RPT" then
	rs("Permissions")=Request("PermissionID")
	elseif Request("PurviewType")="FLW" then
	rs("PermissionsFLW")=Request("PermissionID")
	end if
	rs.update
	rs.close
	set rs=nothing 
	response.write "保存成功，须重新登录才能生效"
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="getUser" then
    GroupID=request("GroupID")
		sql="select a.UserID,b.姓名 from smmsys_PermissionGroupDetails a, "&AllOPENROWSET&" LDERP.dbo.[N-基本资料单头]) b where b.员工代号=a.UserID and a.GroupSnum="&GroupID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		response.write "###"
		while (not rs.eof)
	 		response.write("<option value='"&rs("UserID")&"'>"&rs("UserID")&"/"&rs("姓名")&"</option>")
		  rs.movenext
		wend
		response.write "@@@"
		sql="select a.员工代号 as pk,a.员工代号+'/'+a.姓名 as name,a.部门别 as pspk from [N-基本资料单头] a  where a.离职否='在职' and 职等!='001' and 是否离职=0 "
		sql=sql&" union all  "
		sql=sql&" select distinct c.部门代号 as pk,c.部门名称 as name,'0' as pspk from [N-基本资料单头] a,[G-部门资料表] c  where a.离职否='在职' and a.部门别=c.部门代号 "
		sql=sql&" except "
		sql=sql&" select a.UserID as pk,b.员工代号+'/'+b.姓名 as name,b.部门别 as pspk from  zxpt.dbo.smmsys_PermissionGroupDetails a,[N-基本资料单头] b where b.员工代号=a.UserID and a.GroupSnum="&GroupID
		set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
		response.Write("[")
		do until rs.eof
	%>
	{"SerialNum": "<%=rs("pk")%>", "name": "<%=rs("name")%>", "PSNum": "<%=rs("pspk")%>","val1":"<%=rs("val1")%>","val2":"<%=rs("val2")%>","val3":"<%=rs("val3")%>","val4":"<%=rs("val4")%>"
	<%
		if rs("pspk")="0" then
	%>
	, "isParent":"true"
	<%
		end if
		response.Write("}")
			rs.movenext
			If Not rs.eof Then
			  Response.Write ","
			End If
		loop
		Response.Write "]"
		response.write("###")
		rs.close
		set rs=nothing 
  elseif detailType="Register" then
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
			sql="select * from smmsys_PermissionGroup where SerialNum="&EmpId
			rs.open sql,connzxpt,1,1
			GroupName=rs("GroupName")
			PermissionID=rs("Permissions")

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
      sql="select * from smmsys_PermissionGroup where SerialNum="&EmpId
      rs.open sql,connzxpt,1,1
			PermissionID=rs("PermissionsFLW")
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
  end if
end if
 %>
